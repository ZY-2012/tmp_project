from decimal import Decimal
from pathlib import Path
import json
import zipfile
import pytest
import openpyxl
import aggregate
from aggregate import (money,match_mapping,read_mapping,resolve_month_folder,
                       resolve_mapping_file,consolidate,extract_archive)
from preferences import load_preferences,save_preferences


def salary(path,amount=100):
    path.parent.mkdir(parents=True,exist_ok=True)
    wb=openpyxl.Workbook();ws=wb.active
    ws.title='汇总'
    ws.append(['序号','实发工资','应发工资'])
    ws.append(['合计',amount,110])
    wb.save(path);wb.close()

@pytest.mark.parametrize('text', ['￥ 1,234.56','¥1，234.56','$1,234.56','1\u3000234.56'])
def test_currency(text):
    assert money(text)==Decimal('1234.56')

@pytest.mark.parametrize('text', ['NaN','Infinity',True,'=SUM(A1:A2)'])
def test_invalid_money(text):
    assert money(text) is None


def test_normalized_mapping_and_strict_priority():
    mapping=[('公司(甲)-项目.xls',{'付款组织':'normalized'}),('项目.xlsx',{'付款组织':'strict'})]
    assert match_mapping('姓名-公司（甲）—项目.xlsx',mapping)['付款组织']=='strict'
    assert match_mapping('姓名-公司（甲）—项 目-补充.XLSX',mapping)['付款组织']=='normalized'
    assert match_mapping('其他公司.xlsx',mapping) is None


def test_mapping_title_and_auto_discovery(tmp_path):
    p=tmp_path/'基础信息映射表格.xlsx'
    wb=openpyxl.Workbook();ws=wb.active
    ws.append(['主数据标题']);ws.append(['工 作 簿','付款组织','付款处理'])
    ws.append(['甲.xlsx','演示组织',0]);wb.save(p);wb.close()
    assert resolve_mapping_file(directories=[tmp_path])==p
    assert read_mapping(p)[0][1]['付款处理']==0
    assert resolve_mapping_file(directories=[tmp_path/'no']) is None
    with pytest.raises(FileNotFoundError):
        resolve_mapping_file(tmp_path/'missing.xls')


def test_ambiguous_mapping(tmp_path):
    for name in ['映射甲.xls','映射乙.xlsx']:
        (tmp_path/name).touch()
    with pytest.raises(ValueError,match='多个候选'):
        resolve_mapping_file(directories=[tmp_path])


def test_month_flat_and_variant(tmp_path):
    folder=tmp_path/'2026-10工资数据'
    salary(folder/'子目录'/'甲.xlsx')
    assert resolve_month_folder(tmp_path,2026,10)==folder
    assert resolve_month_folder(folder,2026,10)==folder
    direct=tmp_path/'flat'
    salary(direct/'甲.xlsx')
    assert resolve_month_folder(direct,2026,10)==direct
    with pytest.raises(FileNotFoundError):
        resolve_month_folder(tmp_path,2026,11)


def test_auto_mode_missing_mapping(tmp_path):
    folder=tmp_path/'2026年10月工资表'
    salary(folder/'子目录'/'甲.xlsx')
    result,out=consolidate('2026-10',base_dir=tmp_path,password='')
    assert out==folder/'集团工资汇总_2026-10.xlsx'
    assert result.total==100 and result.matched==0 and result.complete
    assert out.exists() and out.with_suffix('.issues.json').exists()


def test_zip_input_end_to_end(tmp_path):
    source=tmp_path/'sources'
    salary(source/'甲.xlsx',Decimal('1.23'))
    package=tmp_path/'工资.zip'
    with zipfile.ZipFile(package,'w') as z:
        z.write(source/'甲.xlsx','深层/甲.xlsx')
    result,out=consolidate('202610',input_dir=package,password='')
    assert result.files==1 and result.total==Decimal('1.23')
    assert out.parent==tmp_path


def test_partial_and_source_overwrite(tmp_path):
    source=tmp_path/'source'
    p=source/'甲.xlsx'
    salary(p,None)
    with pytest.raises(ValueError,match='非汇总文件'):
        consolidate('2026-10',source,p,password='',overwrite=True)
    result,out=consolidate('2026-10',source,allow_partial=True,password='')
    wb=openpyxl.load_workbook(out)
    assert not result.complete and '不完整' in wb.active['B3'].value
    wb.close()


def test_timestamp_on_locked(tmp_path,monkeypatch):
    source=tmp_path/'src'
    salary(source/'甲.xlsx')
    original=aggregate.write_report
    calls=[]
    def fake(result,output,overwrite=False):
        calls.append(output)
        if len(calls)==1:
            raise PermissionError('locked')
        return original(result,output,overwrite)
    monkeypatch.setattr(aggregate,'write_report',fake)
    result,out=consolidate('2026-10',source,password='',on_locked='timestamp')
    assert len(calls)==2 and out!=calls[0] and out.exists()


def test_preferences_never_save_password(tmp_path):
    p=tmp_path/'prefs.json'
    save_preferences({'base_dir':'示例目录','password':'test-secret','month':'2026-10'},p)
    assert 'password' not in p.read_text(encoding='utf-8')
    assert load_preferences(p)['base_dir']=='示例目录'
    p.write_text('bad json')
    assert load_preferences(p)=={}


def test_zip_link_rejected(tmp_path):
    p=tmp_path/'link.zip'
    with zipfile.ZipFile(p,'w') as z:
        info=zipfile.ZipInfo('link')
        info.external_attr=0o120777 << 16
        z.writestr(info,'target')
    with pytest.raises(ValueError,match='不安全'):
        extract_archive(p,tmp_path/'out')


def test_gui_background_job_and_preview(tmp_path,monkeypatch):
    import os,sys,time
    if sys.platform!='win32' and not os.environ.get('DISPLAY'):
        pytest.skip('GUI needs a display; Windows build runs this test')
    from app import App
    from tkinter import messagebox
    monkeypatch.setattr(messagebox,'showerror',lambda title,text:pytest.fail(text))
    source=tmp_path/'source'
    salary(source/'甲.xlsx',Decimal('88.88'))
    app=App(preferences_path=tmp_path/'prefs.json')
    app.withdraw()
    try:
        app.mode.set('direct');app.update_mode()
        app.source.set(str(source))
        app.values['month'].set('2026-10')
        app.values['output'].set(str(tmp_path/'result.xlsx'))
        app.values['password'].set('test-only-not-saved')
        app.start()
        deadline=time.monotonic()+15
        while app.busy and time.monotonic()<deadline:
            app.update()
            time.sleep(.02)
        assert not app.busy and app.output.exists()
        assert len(app.preview.get_children())==1
        assert app.metrics['total'].get()=='88.88'
        assert app.values['password'].get()==''
        assert 'test-only-not-saved' not in (tmp_path/'prefs.json').read_text(encoding='utf-8')
    finally:
        app.destroy()
