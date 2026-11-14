#!/usr/bin/env python3
"""验证并打包本技能，无第三方依赖，不包含工资数据或应用二进制。"""
from pathlib import Path
import argparse
import hashlib
import re
import zipfile


def validate(root):
    main = root / 'SKILL.md'
    text = main.read_text(encoding='utf-8')
    match = re.match(r'\A---\nname: ([a-z0-9-]+)\ndescription: (.+)\n---\n', text)
    if not match or match[1] != root.name:
        raise ValueError('技能frontmatter缺失，或name与目录名不同')
    files = [main] + sorted((root / 'references').glob('*.md')) + sorted((root / 'scripts').glob('*.py'))
    if len(files) < 3:
        raise ValueError('缺少参考文件或打包脚本')
    for source in files:
        if source.is_symlink():
            raise ValueError(f'不打包链接：{source.name}')
        content = source.read_text(encoding='utf-8')
        if source.suffix == '.md':
            for target in re.findall(r'\]\(([^)]+)\)', content):
                if re.match(r'^[a-zA-Z][a-zA-Z0-9+.-]*:', target) or target.startswith('#'):
                    continue
                resolved = (source.parent / target.split('#')[0]).resolve()
                if not resolved.is_relative_to(root.resolve()) or not resolved.is_file():
                    raise ValueError(f'无效本地引用：{source.name} -> {target}')
    return files


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path, help='输出ZIP路径；不提供时仅验证')
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    files = validate(root)
    if args.output:
        output = args.output.resolve()
        if output.exists():
            parser.error('输出文件已存在，请使用新路径')
        output.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as bundle:
            for source in files:
                bundle.write(source, Path(root.name) / source.relative_to(root))
        with zipfile.ZipFile(output) as bundle:
            if bundle.testzip() is not None or f'{root.name}/SKILL.md' not in bundle.namelist():
                raise ValueError('技能ZIP回读失败')
        digest = hashlib.sha256(output.read_bytes()).hexdigest()
        output.with_suffix(output.suffix + '.sha256').write_text(f'{digest}  {output.name}\n', encoding='utf-8')
        print(f'已验证并打包 {len(files)} 个文件：{output.name}\nSHA256: {digest}')
    else:
        print(f'技能结构、UTF-8内容与本地引用验证通过，共{len(files)}个文件')


if __name__ == '__main__':
    main()
