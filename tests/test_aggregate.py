from decimal import Decimal
import io
import json
from pathlib import Path
import zipfile
import pytest
import openpyxl
import msoffcrypto
from aggregate import (parse_month, parse_summary, inspect_sources, read_mapping, source_files,
                       write_report, consolidate, extract_archive)


def make(path, sheets=('汇总表',), amount=123.45, header='实发工资（汇总）'):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name in sheets:
        ws = wb.create_sheet(name)
        ws.append(['月度工资'])
        ws.append(['序号', '所属部门', header])
        ws.append([1, '项目部', 1])
        ws.append(['合计', None, amount])
        ws.append(['说明：仅供核对'])
    wb.save(path)
    wb.close()


def mapping(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['工作簿','付款组织','付款处理','发放账号'])
    ws.append(['公司.xlsx','集团',0,'00123'])
    ws.append(['甲公司.xlsx','甲公司',0,'00001'])
    ws.append(['甲公司.xlsx','重复不应取',1,'9'])
    wb.save(path)
    wb.close()

@pytest.mark.parametrize('value', ['2026-10','2026年10月','202610'])
def test_month(value):
    assert parse_month(value) == (2026,10)

@pytest.mark.parametrize('value', ['2026-00','2026-13','错误'])
def test_bad_month(value):
    with pytest.raises(ValueError):
        parse_month(value)


def test_multi_sheet_and_mapping(tmp_path):
    source = tmp_path/'人员-2026年10月甲公司.xlsx'
    make(source, ('公司总部','专业公司','工资明细'))
    mp = tmp_path/'映射.xlsx'
    mapping(mp)
    result = inspect_sources([source], read_mapping(mp))
    assert len(result.rows) == 2
    assert result.total == Decimal('246.90')
    assert result.matched == 2
    assert all(r['付款组织'] == '甲公司' and r['付款处理'] == 0 and r['发放账号'] == '00001' for r in result.rows)
    out = tmp_path/'report.xlsx'
    write_report(result,out)
    wb = openpyxl.load_workbook(out)
    assert wb.sheetnames == ['报告']
    assert wb.active['B2'].value == '报告：成功合并 2 个工作表。'
    assert wb.active['D7'].value == 246.9
    assert wb.active['F5'].value == 0
    wb.close()
    with pytest.raises(FileExistsError):
        write_report(result,out)


def test_discovery_exclusions(tmp_path):
    for name in ['~$temp.xlsx','集团工资汇总_2026-10.xlsx','公司.xlsx','映射.xlsx']:
        (tmp_path/name).touch()
    (tmp_path/'子目录').mkdir()
    (tmp_path/'子目录'/'另一公司.XLSX').touch()
    assert len(source_files(tmp_path,tmp_path/'映射.xlsx')) == 2


def test_block_missing_amount_and_preserve_output(tmp_path):
    source = tmp_path/'source'
    source.mkdir()
    make(source/'甲公司.xlsx',amount=None)
    mp = tmp_path/'map.xlsx'
    mapping(mp)
    out = tmp_path/'old.xlsx'
    out.write_bytes(b'original')
    with pytest.raises(ValueError, match='已阻止输出'):
        consolidate('2026-10',source,out,mapping_path=mp,overwrite=True)
    assert out.read_bytes() == b'original'
    assert json.loads(out.with_suffix('.issues.json').read_text(encoding='utf-8'))['金额缺失数'] == 1


def test_last_total_and_header_priority(tmp_path):
    p = tmp_path/'x.xlsx'
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['序号','人数','其他实发工资','实发工资(汇总)'])
    ws.append(['小计',1,999,50])
    ws.append(['合计',1,999,75])
    ws.append(['总计：',1,999,100])
    ws.append(['备注'])
    amount,index,fallback = parse_summary(ws,'实发工资')
    assert (amount,index,fallback) == (Decimal('100'),4,False)


def test_no_total_fallback():
    ws = openpyxl.Workbook().active
    ws.append(['序号','实发工资'])
    ws.append([1,42])
    assert parse_summary(ws,'实发工资') == (Decimal('42'),2,True)


def test_encrypted_and_wrong_password(tmp_path):
    plain = tmp_path/'plain.xlsx'
    encrypted = tmp_path/'甲公司.xlsx'
    make(plain)
    with plain.open('rb') as src, encrypted.open('wb') as dst:
        msoffcrypto.OfficeFile(src).encrypt('test-only-secret',dst)
    result = inspect_sources([encrypted],[],password='test-only-secret')
    assert result.total == Decimal('123.45') and result.complete
    bad = inspect_sources([encrypted],[],password='wrong')
    assert not bad.complete and not bad.rows


def test_zip_and_traversal(tmp_path):
    z = tmp_path/'ok.zip'
    with zipfile.ZipFile(z,'w') as f:
        f.writestr('folder/a.txt','test')
    extract_archive(z,tmp_path/'out')
    assert (tmp_path/'out/folder/a.txt').read_text(encoding='utf-8') == 'test'
    with zipfile.ZipFile(z,'w') as f:
        f.writestr('../escape','bad')
    with pytest.raises(ValueError,match='不安全'):
        extract_archive(z,tmp_path/'out')
    assert not (tmp_path/'escape').exists()


def test_formula_injection_as_text(tmp_path):
    p = tmp_path/'甲公司.xlsx'
    make(p)
    r = inspect_sources([p],[('甲公司.xlsx',{'付款组织':'=1+1'})])
    out = tmp_path/'result.xlsx'
    write_report(r,out)
    wb = openpyxl.load_workbook(out)
    assert wb.active['C5'].data_type == 's'
    wb.close()


def test_empty_does_not_write(tmp_path):
    from aggregate import Result
    with pytest.raises(ValueError):
        write_report(Result(),tmp_path/'empty.xlsx')
    assert not (tmp_path/'empty.xlsx').exists()
