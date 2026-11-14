"""集团月度工资汇总；业务引擎与命令行入口。"""
from __future__ import annotations
import argparse
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import io
import json
import os
from pathlib import Path
import re
import tempfile
import shutil
from datetime import datetime
import zipfile

import msoffcrypto
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

COLUMNS = ['工作簿', '付款组织', '金额', '复核人', '付款处理', '发放账号', '出纳']
FILL_COLUMNS = COLUMNS[1:2] + COLUMNS[3:]
KEYWORDS = ['序号', '所属部门', '人数', '合同工资', '应发工资', '实发工资', '姓名']
CONFIG_PATH = Path.home() / '.workbuddy' / 'salary-consolidation' / 'config.json'

@dataclass
class Result:
    rows: list = field(default_factory=list)
    issues: list = field(default_factory=list)
    files: int = 0
    matched: int = 0
    @property
    def total(self):
        return sum((r['金额'] for r in self.rows if r['金额'] is not None), Decimal('0')).quantize(Decimal('.01'))
    @property
    def missing_amounts(self):
        return sum(r['金额'] is None for r in self.rows)
    @property
    def complete(self):
        return bool(self.rows) and not self.missing_amounts and not any(i['级别'] == '错误' for i in self.issues)


def parse_month(value):
    match = re.fullmatch(r'(\d{4})(?:[-年/.](\d{1,2})月?|(\d{2}))', str(value).strip())
    if not match:
        raise ValueError('月份格式应为 YYYY-MM、YYYY年M月或YYYYMM')
    year, month = int(match[1]), int(match[2] or match[3])
    if not 1 <= month <= 12 or not 1900 <= year <= 9999:
        raise ValueError('月份或年份超出范围')
    return year, month


def normalize_header(value):
    value = re.sub(r'\s+', '', str(value or ''))
    previous = None
    while previous != value:
        previous, value = value, re.sub(r'[（(【\[][^（）()【】\[\]]*[)）】\]]', '', value)
    return value


def money(value):
    if value is None or value == '' or isinstance(value, bool):
        return None
    try:
        number = Decimal(re.sub(r'[,，¥￥$\s\u3000\xa0]', '', str(value)))
        return number if number.is_finite() else None
    except InvalidOperation:
        return None


def load_workbook(path, password=''):
    try:
        return openpyxl.load_workbook(path, data_only=True, read_only=True)
    except (zipfile.BadZipFile, OSError):
        with open(path, 'rb') as source:
            office = msoffcrypto.OfficeFile(source)
            if not password:
                raise ValueError('文件已加密，请输入打开密码')
            office.load_key(password=password, verify_password=True)
            buffer = io.BytesIO()
            office.decrypt(buffer)
        buffer.seek(0)
        return openpyxl.load_workbook(buffer, data_only=True, read_only=True)


def parse_summary(ws, amount_field):
    rows = ws.iter_rows(values_only=True)
    first = []
    for _, row in zip(range(15), rows):
        first.append(row)
    scores = [sum(any(k in normalize_header(c) for k in KEYWORDS) for c in row) for row in first]
    if not scores or max(scores) < 2:
        raise ValueError('前15行未找到表头（至少命中两个关键字）')
    hi = scores.index(max(scores))
    headers = [normalize_header(c) for c in first[hi]]
    exact = [i for i, h in enumerate(headers) if h == amount_field]
    candidates = exact or [i for i, h in enumerate(headers) if amount_field in h]
    if not candidates:
        raise ValueError(f'表头缺少{amount_field}列')
    last = total = None
    total_index = None
    from itertools import chain
    for index, row in enumerate(chain(first[hi + 1:], rows), start=hi + 2):
        if not any(c is not None and c != '' for c in row):
            continue
        last = (index, row)
        if any(k in str(row[0] or '') for k in ['合计', '总计']):
            total, total_index = row, index
    fallback = total is None
    if fallback:
        if last is None:
            raise ValueError('表头后无数据')
        total_index, total = last
    col = candidates[0]
    amount = money(total[col] if col < len(total) else None)
    return amount, total_index, fallback


def read_mapping(path):
    path = Path(path)
    if path.suffix.lower() == '.xls':
        import xlrd
        book = xlrd.open_workbook(str(path))
        sheet = book.sheet_by_index(0)
        rows = [sheet.row_values(i) for i in range(sheet.nrows)]
        book.release_resources()
    else:
        book = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            rows = list(book.worksheets[0].values)
        finally:
            book.close()
    if not rows:
        raise ValueError('映射表为空')
    header_index = next((i for i, row in enumerate(rows[:15]) if any(normalize_header(v) == '工作簿' for v in row)), None)
    if header_index is None:
        raise ValueError('映射表前15行缺少“工作簿”列')
    headers = {normalize_header(v): i for i, v in enumerate(rows[header_index]) if v is not None}
    mapping = {}
    for row in rows[header_index + 1:]:
        def at(name):
            i = headers.get(name)
            return row[i] if i is not None and i < len(row) else None
        key = str(at('工作簿') or '').strip()
        if key and key not in mapping:
            mapping[key] = {c: at(c) for c in FILL_COLUMNS}
    return sorted(mapping.items(), key=lambda pair: len(pair[0]), reverse=True)


def source_files(source, mapping_path=None, output_path=None):
    excluded = {Path(p).resolve() for p in (mapping_path, output_path) if p}
    return sorted((p for p in Path(source).rglob('*') if p.is_file() and p.suffix.lower() == '.xlsx'
                   and not p.name.startswith(('~$', '集团工资汇总', '基础信息映射', '映射表')) and p.resolve() not in excluded), key=lambda p: str(p))


def normalize_workbook_key(value):
    text = re.sub(r'\.xlsx?$', '', str(value).strip(), flags=re.IGNORECASE)
    text = re.sub(r'\s+', '', text)
    return re.sub(r'[—－–―]', '-', text.replace('（', '(').replace('）', ')')).casefold()


def match_mapping(filename, mapping):
    # 严格包含优先；仅无严格匹配时做名称归一化，和原工具一致。
    matches = [(key, info) for key, info in mapping if key in filename]
    if matches:
        return max(matches, key=lambda pair: len(pair[0]))[1]
    target = normalize_workbook_key(filename)
    matches = [(normalize_workbook_key(key), info) for key, info in mapping
               if normalize_workbook_key(key) and normalize_workbook_key(key) in target]
    return max(matches, key=lambda pair: len(pair[0]))[1] if matches else None


def resolve_month_folder(base, year, month):
    base = Path(base)
    if not base.is_dir():
        raise FileNotFoundError(f'工资根目录不存在：{base}')
    exact = base / f'{year}年{month}月工资表'
    if exact.is_dir():
        return exact
    markers = [f'{year}年{month}月', f'{year}年{month:02d}月', f'{year}-{month:02d}']
    candidates = sorted(p for p in base.iterdir() if p.is_dir() and any(marker in p.name for marker in markers))
    if len(candidates) == 1:
        return candidates[0]
    if len(candidates) > 1:
        raise ValueError('找到多个当月目录，请直接选择源目录')
    if source_files(base) and any(p.suffix.lower() == '.xlsx' and not p.name.startswith(('~$', '集团工资汇总', '基础信息映射', '映射表')) for p in base.iterdir() if p.is_file()):
        return base
    if any(p.is_file() and p.suffix.lower() in ('.zip', '.rar') for p in base.iterdir()):
        return base
    # 允许直接选择“当月目录”，其中源表放在任意层子目录。
    if any(marker in base.name for marker in markers) and source_files(base):
        return base
    raise FileNotFoundError(f'未找到{year}年{month}月工资表；请直接选择当月目录或压缩包')


def resolve_mapping_file(explicit=None, directories=()):
    if explicit:
        path = Path(explicit)
        if not path.is_file():
            raise FileNotFoundError(f'指定映射表不存在：{path}')
        return path
    roots = list(dict.fromkeys(Path(p).resolve() for p in directories if p))
    for root in roots:
        for name in ('基础信息映射表格.xls', '基础信息映射表格.xlsx'):
            if (root / name).is_file():
                return root / name
    candidates = sorted({p for root in roots if root.is_dir() for p in root.iterdir()
                         if p.is_file() and p.suffix.lower() in ('.xls', '.xlsx')
                         and not p.name.startswith('~$') and ('映射' in p.name or '基础信息' in p.name)})
    if len(candidates) == 1:
        return candidates[0]
    if len(candidates) > 1:
        raise ValueError('发现多个候选映射表，请明确选择其中一份')
    return None


def inspect_sources(files, mapping, password='', amount_field='实发工资', progress=None):
    result = Result(files=len(files))
    def issue(level, path, sheet, reason):
        result.issues.append({'级别': level, '工作簿': path.name, '工作表': sheet, '说明': reason})
    for index, path in enumerate(files, 1):
        if progress:
            progress(index, len(files), path.name)
        book = None
        try:
            book = load_workbook(path, password)
            sheets = [s for s in book.sheetnames if '明细' not in s]
            if not sheets:
                issue('错误', path, '', '无非明细工作表')
            for name in sheets:
                try:
                    amount, row_index, fallback = parse_summary(book[name], amount_field)
                    row = {c: None for c in COLUMNS}
                    row.update({'工作簿': path.name, '金额': amount, '工作表': name, '源合计行': row_index})
                    match = match_mapping(path.name, mapping)
                    if match is not None:
                        row.update(match)
                        result.matched += 1
                        if not any(key in path.name for key, _ in mapping):
                            issue('提示', path, name, '通过名称归一化匹配付款信息（可能忽略扩展名/项目后缀），请人工复核')
                    else:
                        issue('提示', path, name, '未匹配付款信息，请补充映射表')
                    if amount is None:
                        issue('错误', path, name, '金额为空或非数字；若为公式，请先在Excel计算并保存')
                    if fallback:
                        issue('提示', path, name, f'未找到合计标记，使用最后非空行（第{row_index}行），请核对')
                    result.rows.append(row)
                except Exception as error:
                    issue('错误', path, name, str(error))
        except Exception as error:
            # 不在日志中写密码或任何解密内容。
            message = '密码错误或文件无法解密' if 'key' in type(error).__name__.lower() else str(error)
            issue('错误', path, '', message)
        finally:
            if book is not None:
                book.close()
    return result


def safe_cell(cell, value):
    cell.value = value
    if isinstance(value, str):
        cell.data_type = 's'  # 文件名/映射文本不能作为Excel公式执行


def write_report(result, output, overwrite=False):
    if not result.rows:
        raise ValueError('没有可输出的汇总行；未生成或覆盖文件')
    output = Path(output)
    if output.exists() and not overwrite:
        raise FileExistsError('输出文件已存在，请更换名称或明确允许覆盖')
    book = openpyxl.Workbook()
    ws = book.active
    ws.title = '报告'
    ws['B2'] = f'报告：成功合并 {len(result.rows)} 个工作表。'
    if not result.complete:
        ws['B3'] = '注意：存在读取错误或金额缺失，本报告不完整，请查阅异常清单。'
    for col, title in enumerate(COLUMNS, 2):
        ws.cell(4, col, title)
    for ri, row in enumerate(result.rows, 5):
        for ci, key in enumerate(COLUMNS, 2):
            value = row[key]
            if isinstance(value, Decimal):
                value = float(value)
            safe_cell(ws.cell(ri, ci), value)
    end = len(result.rows) + 5
    ws.cell(end, 2, '合计')
    ws.cell(end, 4, float(result.total))
    ws.freeze_panes = 'D5'
    ws.auto_filter.ref = f'B4:H{end-1}'
    for col, width in zip('BCDEFGH', [85, 42, 20, 18, 22, 26, 18]):
        ws.column_dimensions[col].width = width
    for row in ws.iter_rows(min_row=4, max_row=end, min_col=2, max_col=8):
        for cell in row:
            cell.alignment = Alignment(vertical='center')
            if cell.row in (4, end):
                cell.fill = PatternFill('solid', fgColor='163B61')
                cell.font = Font(color='FFFFFF', bold=True)
            elif cell.row % 2:
                cell.fill = PatternFill('solid', fgColor='EFF5FA')
        row[2].number_format = '#,##0.00'
    ws.sheet_view.showGridLines = False
    output.parent.mkdir(parents=True, exist_ok=True)
    # 原子替换：保存失败时原成品保持完整。
    fd, temp = tempfile.mkstemp(suffix='.xlsx', dir=output.parent)
    os.close(fd)
    try:
        book.save(temp)
        if output.exists() and not overwrite:
            raise FileExistsError('输出文件已存在')
        os.replace(temp, output)
    finally:
        Path(temp).unlink(missing_ok=True)
        book.close()
    return output


def write_issues(result, path):
    # 单独日志，保持成品仅一个“报告”sheet。诊断含公司名，仅保存本地。
    Path(path).write_text(json.dumps({'文件数': result.files, '汇总行数': len(result.rows),
        '完整': result.complete, '金额合计': str(result.total), '金额缺失数': result.missing_amounts,
        '映射匹配数': result.matched, '异常': result.issues}, ensure_ascii=False, indent=2), encoding='utf-8')


def get_password():
    value = os.environ.get('SALARY_XLSX_PASSWORD', '')
    if value:
        return value
    if CONFIG_PATH.exists():
        return str(json.loads(CONFIG_PATH.read_text(encoding='utf-8')).get('password', ''))
    return ''


def configure_rar_tools():
    import rarfile
    # 安装在默认位置的WinRAR/7-Zip无需手动配置PATH；Windows自带tar为后备。
    roots = [Path(os.environ.get(key, fallback)) for key, fallback in
             [('ProgramFiles', r'C:\Program Files'), ('ProgramFiles(x86)', r'C:\Program Files (x86)')]]
    for attribute, relative in [('UNRAR_TOOL', 'WinRAR/UnRAR.exe'), ('SEVENZIP_TOOL', '7-Zip/7z.exe')]:
        if not shutil.which(getattr(rarfile, attribute)):
            found = next((root / relative for root in roots if (root / relative).is_file()), None)
            if found:
                setattr(rarfile, attribute, str(found))
    if os.name == 'nt':
        tar = Path(os.environ.get('SystemRoot', r'C:\Windows')) / 'System32/tar.exe'
        if tar.is_file():
            rarfile.BSDTAR_TOOL = str(tar)


def extract_archive(path, destination, password=None):
    """仅提取普通文件，禁止路径穿越/链接；RAR依赖系统解压后端。"""
    path, destination = Path(path), Path(destination).resolve()
    if path.suffix.lower() == '.zip':
        archive = zipfile.ZipFile(path)
        entries = [(i.filename, i.is_dir(), (i.external_attr >> 16) & 0o170000 == 0o120000, i.file_size, i) for i in archive.infolist()]
    elif path.suffix.lower() == '.rar':
        import rarfile
        configure_rar_tools()
        archive = rarfile.RarFile(path)
        if password:
            archive.setpassword(password)
        entries = [(i.filename, i.isdir(), i.is_symlink(), i.file_size, i) for i in archive.infolist()]
    else:
        raise ValueError('仅支持ZIP或RAR压缩包')
    try:
        if sum(e[3] for e in entries) > 2 * 1024**3 or len(entries) > 10000:
            raise ValueError('压缩包展开超过2GB或10000项限制，请手动整理')
        for name, is_dir, is_link, _, entry in entries:
            normalized = name.replace('\\', '/')
            target = (destination / normalized).resolve()
            if is_link or not target.is_relative_to(destination) or ':' in normalized:
                raise ValueError('压缩包含不安全路径或链接')
        for name, is_dir, _, _, entry in entries:
            target = destination / name.replace('\\', '/')
            if is_dir:
                target.mkdir(parents=True, exist_ok=True)
            else:
                target.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(entry, pwd=password.encode('utf-8') if password and path.suffix.lower() == '.zip' else password) as src, target.open('wb') as dst:
                    import shutil
                    shutil.copyfileobj(src, dst)
    finally:
        archive.close()


def consolidate(month, input_dir=None, output_path=None, amount_field='实发工资', mapping_path=None,
                base_dir='.', password=None, overwrite=False, allow_partial=False, progress=None,
                on_locked='error', log=None):
    year, mon = parse_month(month)
    if amount_field not in ('实发工资', '应发工资'):
        raise ValueError('金额口径仅支持实发工资或应发工资')
    notify = log or (lambda text: None)
    base = Path(base_dir)
    source = Path(input_dir) if input_dir else resolve_month_folder(base, year, mon)
    if not source.exists():
        raise FileNotFoundError(f'源目录或压缩包不存在：{source}')
    folder = source.parent if source.is_file() or '发放工资表' in source.name or 'to胡经理' in source.name else source
    mapping = resolve_mapping_file(mapping_path, [base, source if source.is_dir() else source.parent, folder, folder.parent])
    master = read_mapping(mapping) if mapping else []
    notify(f'月份：{year}-{mon:02d} | 输入：{source}')
    notify(f'映射表：{mapping}' if mapping else '未找到映射表，付款信息将留空，请在结果中核对。')
    output = Path(output_path) if output_path else folder / f'集团工资汇总_{year}-{mon:02d}.xlsx'
    if output.suffix.lower() != '.xlsx':
        raise ValueError('输出文件必须使用.xlsx扩展名')
    if output.resolve() == (mapping.resolve() if mapping else None):
        raise ValueError('输出路径不能覆盖映射表')
    if output.exists() and not overwrite:
        raise FileExistsError(f'成品已存在：{output}；请选择新路径或勾选允许覆盖')
    if output.exists() and source.is_file() and output.resolve() == source.resolve():
        raise ValueError('输出路径不能覆盖输入文件')
    if output.exists() and source.is_dir() and output.resolve().is_relative_to(source.resolve()) and not output.name.startswith('集团工资汇总'):
        existing = None
        try:
            existing = openpyxl.load_workbook(output, read_only=True, data_only=True)
            is_report = existing.sheetnames == ['报告'] and [existing.active.cell(4, c).value for c in range(2, 9)] == COLUMNS
        except Exception:
            is_report = False
        finally:
            if existing is not None:
                existing.close()
        if not is_report:
            raise ValueError('输出路径指向源目录中的非汇总文件，已阻止覆盖')
    pw = get_password() if password is None else password
    with tempfile.TemporaryDirectory(prefix='salary-') as temp:
        if source.is_file():
            notify('正在临时解压压缩包…')
            extract_archive(source, temp, pw or None)
            files = source_files(temp, mapping, output)
        else:
            files = source_files(source, mapping, output)
            if not files:
                archives = sorted(p for p in source.rglob('*') if p.suffix.lower() in ('.rar', '.zip'))
                for index, archive in enumerate(archives):
                    notify(f'正在临时解压：{archive.name}')
                    extract_archive(archive, Path(temp) / str(index), pw or None)
                files = source_files(temp, mapping, output)
        if not files:
            raise ValueError('未找到源xlsx文件（已排除旧汇总、映射表、临时文件）')
        if output.resolve() in {f.resolve() for f in files}:
            raise ValueError('输出路径不能覆盖源工资表')
        result = inspect_sources(files, master, pw, amount_field, progress)
    output.parent.mkdir(parents=True, exist_ok=True)
    if not result.complete and not allow_partial:
        diagnostic = output.with_suffix('.issues.json')
        write_issues(result, diagnostic)
        raise ValueError(f'存在读取错误或金额缺失，已阻止输出；详情见 {diagnostic}')
    try:
        write_report(result, output, overwrite)
    except PermissionError:
        if on_locked != 'timestamp':
            raise PermissionError('无法写入成品，请关闭Excel中的目标文件或更换输出目录') from None
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
        output = output.with_name(f'{output.stem}_{stamp}{output.suffix}')
        notify(f'原目标文件无法写入，已使用新文件名：{output.name}')
        write_report(result, output, overwrite=False)
    write_issues(result, output.with_suffix('.issues.json'))
    notify(f'完成：源文件{result.files}个 / 汇总{len(result.rows)}行 / 金额缺失{result.missing_amounts}行')
    return result, output


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--month', required=True)
    parser.add_argument('--input')
    parser.add_argument('--output')
    parser.add_argument('--mapping')
    parser.add_argument('--base-dir', default='.')
    parser.add_argument('--amount-field', choices=['实发工资', '应发工资'], default='实发工资')
    parser.add_argument('--on-locked', choices=['error', 'timestamp'], default='error')
    parser.add_argument('--overwrite', action='store_true')
    parser.add_argument('--allow-partial', action='store_true', help='明确允许生成不完整报告')
    args = parser.parse_args()
    try:
        result, output = consolidate(args.month, args.input, args.output, args.amount_field, args.mapping,
                                     args.base_dir, overwrite=args.overwrite, allow_partial=args.allow_partial, on_locked=args.on_locked)
        print(f'已保存：{output}\n文件 {result.files} | 汇总行 {len(result.rows)} | 合计 {result.total} | 映射 {result.matched}/{len(result.rows)}')
        print(f'异常/提示 {len(result.issues)} 条；详见同名 .issues.json')
    except Exception as error:
        parser.exit(1, f'处理失败：{error}\n')

if __name__ == '__main__':
    main()
