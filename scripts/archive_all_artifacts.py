"""Archive this project's surviving local outputs without changing source files.

Run with Python 3.10+. Copies regular files and verifies SHA-256; dependency
symlinks are documented but never followed. The archive is local-only.
"""
from pathlib import Path
import hashlib
import json
import shutil
import zipfile

PROJECT = Path(__file__).resolve().parent.parent
ARCHIVE = PROJECT / 'archive-local'
THREAD = '01a104be-b43a-7c61-8907-9c0db0fd5cfc'
HOME = Path.home()
OUTPUTS = HOME / 'outputs' / THREAD
GENERATED = HOME / '.codex/generated_images' / THREAD
ZIP_NAMES = [
    '单品利润测算样品包.zip',
    '单品利润测算样品包_v0.2修正版.zip',
    '单品利润测算样品包_v0.3修正版.zip',
    '三位商家AI模拟试填结果.zip',
    '网店老板_v0.3模拟试填.zip',
    '单品成本与退货贡献测算_v0.3_交付包.zip',
]

def sha(path):
    digest = hashlib.sha256()
    with path.open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()

def copy_file(source, destination):
    expected = sha(source)
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists() and sha(destination) != expected:
        raise RuntimeError(f'Existing archive differs; preserve and review: {destination}')
    if not destination.exists():
        shutil.copy2(source, destination)
    if sha(destination) != expected:
        raise RuntimeError(f'Copy verification failed: {source}')
    return {
        'source': str(source),
        'destination': str(destination.relative_to(PROJECT)),
        'bytes': source.stat().st_size,
        'sha256': expected,
    }

def main():
    records, skipped, missing = [], [], []
    for source_root, folder in [(OUTPUTS, '历史产出'), (GENERATED, '图像生成原件')]:
        if not source_root.exists():
            missing.append(str(source_root))
            continue
        for source in sorted(source_root.rglob('*')):
            if source.is_symlink():
                skipped.append({'source': str(source), 'target': str(source.readlink()),
                                'reason': 'Runtime dependency link; not a project artifact'})
            elif source.is_file():
                records.append(copy_file(source, ARCHIVE / folder / source.relative_to(source_root)))
    for name in ZIP_NAMES:
        source = HOME / 'Downloads' / name
        if source.exists():
            with zipfile.ZipFile(source) as package:
                bad = package.testzip()
                if bad is not None:
                    raise RuntimeError(f'ZIP integrity failed: {source}: {bad}')
            record = copy_file(source, ARCHIVE / '原始与历版压缩包' / name)
            record['zip_integrity_passed'] = True
            records.append(record)
        else:
            missing.append(str(source))
    source = Path('/tmp/prepare_shop_project.py')
    if source.is_file():
        records.append(copy_file(source, ARCHIVE / '早期项目整理脚本' / source.name))
    else:
        missing.append(str(source))
    manifest = {
        'prepared_on': '2026-10-06', 'local_only': True,
        'source_files_preserved': True, 'all_copies_sha256_verified': True,
        'copied_file_count': len(records), 'copied_bytes': sum(r['bytes'] for r in records),
        'files': records, 'excluded_dependency_links': skipped,
        'missing_sources': missing,
        'scope': 'All surviving regular files in thread output and image directories, '
                 'six named Downloads ZIPs, and the surviving early preparation script. '
                 'Chat transcripts, deleted temporary files and third-party dependencies are outside scope.',
    }
    ARCHIVE.mkdir(parents=True, exist_ok=True)
    (ARCHIVE / '来源与校验清单.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + '\n')
    inventory_path = PROJECT / '全项目产物索引.json'
    inventory = []
    for file in sorted(PROJECT.rglob('*')):
        if file.is_symlink() or not file.is_file() or file == inventory_path:
            continue
        relative = file.relative_to(PROJECT)
        if any(part in ['.git', 'node_modules', '__pycache__'] for part in relative.parts) or file.name == '.DS_Store':
            continue
        inventory.append({'path': str(relative), 'bytes': file.stat().st_size, 'sha256': sha(file)})
    inventory_path.write_text(json.dumps({
        'prepared_on': '2026-10-06', 'file_count': len(inventory),
        'note': 'Includes current work and local archives; excludes this index itself, OS metadata and dependencies.',
        'files': inventory,
    }, ensure_ascii=False, indent=2) + '\n')
    print(json.dumps({k: v for k, v in manifest.items() if k not in ['files', 'excluded_dependency_links', 'scope']}, ensure_ascii=False))

if __name__ == '__main__':
    main()
