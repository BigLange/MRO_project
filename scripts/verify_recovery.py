"""Portable checks requiring Python's standard library only; NOT a C#/WPF build."""
from pathlib import Path
import hashlib
import json
import re
import sys
from zipfile import ZipFile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
P = '{http://schemas.microsoft.com/winfx/2006/xaml/presentation}'
X = '{http://schemas.microsoft.com/winfx/2006/xaml}'


def main():
    recovery = json.loads((ROOT / 'docs/ui-recovery.json').read_text(encoding='utf-8'))
    manifest = json.loads((ROOT / 'docs/resource-manifest.json').read_text(encoding='utf-8'))
    for item in manifest:
        path = ROOT / item['path']
        assert path.is_file(), f'Missing asset: {path}'
        assert hashlib.sha256(path.read_bytes()).hexdigest() == item['sha256'], f'Asset changed: {path}'
    shared = (ROOT / 'apps/Shared/WindowSupport.cs').read_text(encoding='utf-8')
    for report in recovery:
        project = ROOT / 'apps' / report['app']
        root = ET.parse(project / 'MainWindow.xaml').getroot()
        assert root.tag == P + 'Window'
        assert root.get(X + 'Class') == report['app'] + '.MainWindow'
        code = (project / 'MainWindow.xaml.cs').read_text(encoding='utf-8')
        keys = {node.get(X + 'Key') for node in root.iter() if node.get(X + 'Key')}
        assert set(report['resource_keys']) == keys, 'Resource dictionary keys changed'
        for node in root.iter():
            for key, value in node.attrib.items():
                if key in ('Icon', 'Source') and value.startswith('Assets/'):
                    assert (project / value).is_file(), f'Unresolved asset: {value}'
                if value.startswith('{StaticResource '):
                    assert value[len('{StaticResource '):-1] in keys, f'Unresolved static resource: {value}'
                if key == 'Click':
                    assert re.search(r'void\s+' + re.escape(value) + r'\s*\(', code), f'Missing handler: {value}'
            if node.tag == P + 'Binding':
                binding = node.get('Path')
                assert re.search(r'\b' + re.escape(binding) + r'\s*(?:\{|=>|[,)=])', code + '\n' + shared), f'Missing binding property: {binding}'
        actual_elements = sum('.' not in node.tag.split('}')[-1] for node in root.iter())
        assert actual_elements == report['record_type_counts']['3'] + report.get('additional_elements', 0), 'BAML element count changed'
        print(f"PASS {report['app']}: XML, {actual_elements} elements, resources, bindings, event handlers")
    for project in ROOT.rglob('*.csproj'):
        xml = ET.parse(project).getroot()
        for reference in xml.iter('ProjectReference'):
            assert (project.parent / reference.get('Include')).resolve().is_file(), 'Missing project reference'
    assert len(list(ROOT.rglob('*.csproj'))) == 4
    ET.parse(ROOT / 'Directory.Build.props')
    ET.parse(ROOT / 'apps/Mro.MainApp/App.xaml')
    ET.parse(ROOT / 'apps/Mro.HelperApp/App.xaml')
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    with ZipFile(ROOT / 'SharedContent/ReportTemplates/MRO-ReportTemplate.xlsx') as archive:
        assert not any('externalLink' in name for name in archive.namelist()), 'External link retained in blank template'
        for i in range(1, 5):
            sheet = ET.fromstring(archive.read(f'xl/worksheets/sheet{i}.xml'))
            assert not list(sheet.iter(ns + 'f')), 'Formula retained in blank template'
        detail = ET.fromstring(archive.read('xl/worksheets/sheet2.xml'))
        cells = {cell.get('r'): cell for cell in detail.iter(ns + 'c')}
        for column in ('A', 'C', 'E', 'G'):
            for row in range(3, 28):
                cell = cells.get(f'{column}{row}')
                assert cell is None or not ''.join(cell.itertext()).strip(), 'Historical sample value retained'
    print('PASS blank report template: 100 cleared sample slots, 4 sheets, no formulas/external links')
    print(f'PASS {len(manifest)} resource/template/workbook hashes and 4 project references')
    print('C# compilation, console specs, and WPF rendering have NOT been run by this script.')


if __name__ == '__main__':
    try:
        main()
    except (AssertionError, OSError, ValueError) as error:
        print(f'FAIL {error}', file=sys.stderr)
        sys.exit(1)
