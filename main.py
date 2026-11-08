import sys, json
from pathlib import Path
from PySide6.QtCore import Qt, QUrl, QSettings, QTimer
from PySide6.QtWidgets import (QApplication,QMainWindow,QWidget,QVBoxLayout,QHBoxLayout,QLabel,QPushButton,QFileDialog,QTableWidget,QTableWidgetItem,QTabWidget,QLineEdit,QMessageBox,QFormLayout,QDialog,QDialogButtonBox,QHeaderView,QProgressBar,QCheckBox)
from PySide6.QtWebEngineWidgets import QWebEngineView
from core.excel_reader import read_workbook
from core.mapping import DEFAULT_MAPPING
from core.automation import build_employee_plan, OperationLog

MAIN_URL = "https://kara.tehran115.ir/Shift/YearlyShiftassignment.aspx"

class MappingDialog(QDialog):
    def __init__(self, mapping, parent=None):
        super().__init__(parent); self.setWindowTitle("تنظیم نگاشت کدهای شیفت"); self.edits={}; form=QFormLayout(self)
        labels={"ص":"صبح‌کار","ع":"عصرکار","م":"مرخصی","ن":"نقدی","خالی":"خانه خالی / آف","ض":"کد ض","غ":"کد غ"}
        for key,label in labels.items():
            edit=QLineEdit(mapping.get(key,"")); self.edits[key]=edit; form.addRow(label,edit)
        buttons=QDialogButtonBox(QDialogButtonBox.StandardButton.Save|QDialogButtonBox.StandardButton.Cancel); buttons.accepted.connect(self.accept); buttons.rejected.connect(self.reject); form.addRow(buttons)
    def values(self): return {k:v.text().strip().upper() for k,v in self.edits.items()}

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__(); self.setWindowTitle("Kara Shift Manager | ثبت هوشمند شیفت"); self.resize(1450,950)
        self.data=None; self.queue=[]; self.index=0; self.busy=False; self.awaiting_manual_confirm=False; self._fill_started=False; self._verify_started=False
        self.log=OperationLog(); self.settings=QSettings("PardikTeam","KaraShiftManager")
        saved=self.settings.value("shift_mapping", None); self.mapping=dict(saved) if isinstance(saved, dict) else dict(DEFAULT_MAPPING)
        tabs=QTabWidget(); self.setCentralWidget(tabs)
        page=QWidget(); lay=QVBoxLayout(page)
        title=QLabel("مدیریت و ثبت گروهی برنامه شیفت پرسنل"); title.setStyleSheet("font-size:23px;font-weight:bold;color:#173b70;padding:10px"); lay.addWidget(title)
        bar=QHBoxLayout(); choose=QPushButton("انتخاب فایل Excel"); choose.clicked.connect(self.load_excel); mapbtn=QPushButton("ویرایش نگاشت کدها"); mapbtn.clicked.connect(self.edit_mapping); bar.addWidget(choose); bar.addWidget(mapbtn); self.path_label=QLabel("فایلی انتخاب نشده است"); bar.addWidget(self.path_label,1); lay.addLayout(bar)
        self.summary=QLabel("ابتدا فایل XLSX یا XLSM را انتخاب کنید."); self.summary.setStyleSheet("background:#eaf2fb;padding:12px"); lay.addWidget(self.summary)
        self.table=QTableWidget(0,7); self.table.setHorizontalHeaderLabels(["شیت","ردیف","نام","کد ملی","نوع شیفت","روزها","اعتبارسنجی"]); self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch); self.table.setAlternatingRowColors(True); lay.addWidget(self.table,1)
        actions=QHBoxLayout(); self.start_btn=QPushButton("شروع ثبت واقعی شیفت‌ها"); self.start_btn.clicked.connect(self.start_registration); self.stop_btn=QPushButton("توقف امن پس از ردیف جاری"); self.stop_btn.clicked.connect(self.request_stop); self.stop_btn.setEnabled(False); self.skip_done=QCheckBox("ردیف‌های قبلاً موفق ثبت‌شده را رد کن"); self.skip_done.setChecked(True); actions.addWidget(self.start_btn); actions.addWidget(self.stop_btn); actions.addWidget(self.skip_done); lay.addLayout(actions)
        self.progress=QProgressBar(); self.progress.setValue(0); lay.addWidget(self.progress); self.status=QLabel("آماده — ابتدا در تب سامانه وارد حساب کاربری شوید."); self.status.setWordWrap(True); lay.addWidget(self.status); tabs.addTab(page,"۱. فایل اکسل و ثبت شیفت")
        browser_page=QWidget(); bl=QVBoxLayout(browser_page); row=QHBoxLayout(); openbtn=QPushButton("ورود به سامانه کارا"); openbtn.clicked.connect(lambda:self.browser.setUrl(QUrl("https://kara.tehran115.ir/Login.aspx"))); shiftbtn=QPushButton("باز کردن صفحه اختصاص گروهی سالیانه"); shiftbtn.clicked.connect(lambda:self.browser.setUrl(QUrl(MAIN_URL))); inspectbtn=QPushButton("بررسی وضعیت صفحه"); inspectbtn.clicked.connect(self.inspect_page); row.addWidget(openbtn); row.addWidget(shiftbtn); row.addWidget(inspectbtn); bl.addLayout(row)
        tip=QLabel("۱) وارد سامانه شوید. ۲) صفحه «اختصاص گروهی سالیانه» را باز کنید. ۳) به تب اول برگردید و ثبت را شروع کنید. برنامه پنجره انتخاب پرسنل را باز می‌کند؛ اگر تأیید خودکار قابل تشخیص نباشد، از شما می‌خواهد دکمه سبز تأیید را دستی بزنید. رمز عبور خوانده یا ذخیره نمی‌شود."); tip.setWordWrap(True); tip.setStyleSheet("background:#fff2d8;padding:10px"); bl.addWidget(tip)
        self.browser=QWebEngineView(); self.browser.setUrl(QUrl("https://kara.tehran115.ir/Login.aspx")); bl.addWidget(self.browser,1); tabs.addTab(browser_page,"۲. مرورگر سامانه کارا")
        settings=QWidget(); sl=QVBoxLayout(settings); sl.addWidget(QLabel("نگاشت کدهای اکسل به کدهای سامانه")); self.map_table=QTableWidget(7,3); self.map_table.setHorizontalHeaderLabels(["کد اکسل","مفهوم","کد سامانه"]); self.map_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch); sl.addWidget(self.map_table); self.refresh_map(); sl.addWidget(QLabel("کدهای ض و غ با مقدار ERROR مانع ثبت همان ردیف می‌شوند. مقدار ن و خانه خالی برابر OFF است. برای کد ملی تکراری، فقط ردیف «اصلی» خودکار ثبت می‌شود؛ ردیف «دوم» برای ثبت دستی کنار گذاشته می‌شود.")); tabs.addTab(settings,"۳. تنظیمات")
        self.setStyleSheet("QMainWindow{background:#f4f7fb} QWidget{font-family:Segoe UI;font-size:10pt} QPushButton{background:#245a98;color:white;border:0;border-radius:6px;padding:9px 14px} QPushButton:disabled{background:#9aa8b8} QLineEdit{background:white;border:1px solid #cbd5e1;border-radius:5px;padding:8px} QTableWidget{background:white;gridline-color:#e2e8f0}")

    def load_excel(self):
        path,_=QFileDialog.getOpenFileName(self,"انتخاب برنامه شیفت","","Excel (*.xlsx *.xlsm)")
        if not path:return
        try:
            self.data=read_workbook(path); es=self.data["employees"]; self.table.setRowCount(len(es))
            for i,e in enumerate(es):
                _,errs=build_employee_plan(e,self.mapping); vals=[e["sheet"],str(e["row"]),e["first"]+" "+e["last"],e["nid"],e["shift"],str(len(e["days"])),("خطا: "+"؛ ".join(errs[:2])) if errs else "آماده"]
                for j,val in enumerate(vals): self.table.setItem(i,j,QTableWidgetItem(val))
            counts={}
            for e in es: counts[e["nid"]]=counts.get(e["nid"],0)+1
            repeated=sum(1 for n in counts.values() if n>1)
            self.summary.setText(f"ماه: {self.data['month']} {self.data['year']} | ردیف‌ها: {len(es)} | کد ملی یکتا: {len(counts)} | کد ملی تکراری: {repeated} | شیت عملیاتی: {len(self.data['sheets'])}")
            self.path_label.setText(path); self.status.setText("فایل خوانده شد. برای ثبت، ابتدا وارد سامانه شوید و صفحه اختصاص گروهی سالیانه را باز کنید.")
            if self.data["warnings"]: QMessageBox.warning(self,"هشدار فایل","\n".join(self.data["warnings"]))
        except Exception as exc: QMessageBox.critical(self,"خطا",str(exc))

    def edit_mapping(self):
        dlg=MappingDialog(self.mapping,self)
        if dlg.exec(): self.mapping=dlg.values(); self.settings.setValue("shift_mapping",self.mapping); self.settings.sync(); self.refresh_map();
        if self.data: self.load_table_validation()
    def load_table_validation(self):
        for i,e in enumerate(self.data["employees"]):
            _,errs=build_employee_plan(e,self.mapping); self.table.setItem(i,6,QTableWidgetItem(("خطا: "+"؛ ".join(errs[:2])) if errs else "آماده"))
    def refresh_map(self):
        meanings={"ص":"صبح‌کار","ع":"عصرکار","م":"مرخصی","ن":"نقدی","خالی":"آف / خالی","ض":"کد ض","غ":"کد غ"}
        for i,(k,v) in enumerate(self.mapping.items()):
            for j,val in enumerate((k,meanings.get(k,k),v)):
                if hasattr(self,"map_table"): self.map_table.setItem(i,j,QTableWidgetItem(val))

    def js(self, script, callback): self.browser.page().runJavaScript(script, callback)
    def page_ready(self):
        return "location.href.includes('YearlyShiftassignment.aspx')"
    def start_registration(self):
        if self.busy:return
        if not self.data: QMessageBox.warning(self,"فایل اکسل","ابتدا فایل اکسل را انتخاب کنید."); return
        self.js("location.href", self._start_after_url)
    def _start_after_url(self,url):
        if not url or "YearlyShiftassignment.aspx" not in url:
            QMessageBox.warning(self,"صفحه کارا","ابتدا وارد سامانه شوید و صفحه «اختصاص گروهی سالیانه» را باز کنید."); return
        valid=[]; skipped=[]
        for e in self.data["employees"]:
            _,errs=build_employee_plan(e,self.mapping)
            if errs:
                self.log.append(e,"خطای نگاشت","؛ ".join(errs)); skipped.append(f"{e['first']} {e['last']} ({e['nid']}): {'؛ '.join(errs)}")
            elif len(e["nid"])!=10:
                self.log.append(e,"خطای کد ملی","کد ملی باید دقیقاً ۱۰ رقم باشد"); skipped.append(f"{e['first']} {e['last']}: کد ملی نامعتبر")
            else: valid.append(e)
        if not valid: QMessageBox.warning(self,"ثبت ناممکن","هیچ ردیف معتبری برای ثبت وجود ندارد. موارد ERROR را در نگاشت اصلاح کنید."); return
        # User decision: duplicate-NID (two-shift) employees are auto-registered
        # only once, using the row explicitly marked as primary. Secondary rows
        # are intentionally left for the user to enter manually.
        grouped={}
        for employee in valid:
            grouped.setdefault(employee['nid'], []).append(employee)
        auto_rows=[]
        manual_secondary=[]
        for nid, rows in grouped.items():
            if len(rows)==1:
                auto_rows.append(rows[0])
                continue
            primary=[r for r in rows if 'اصلی' in (r.get('shift') or '').replace('ي','ی').replace('ك','ک')]
            if len(primary)==1:
                auto_rows.append(primary[0])
                manual_secondary.extend(r for r in rows if r is not primary[0])
                for r in rows:
                    if r is not primary[0]:
                        self.log.append(r,'ثبت دستی','پرسنل دوشیفته است؛ طبق تصمیم کاربر فقط ردیف اصلی خودکار ثبت می‌شود.')
            else:
                for r in rows:
                    self.log.append(r,'نیازمند بررسی','کد ملی تکراری است اما ردیف اصلی با اطمینان تشخیص داده نشد؛ خودکار ثبت نشد.')
                skipped.append(f'{rows[0]["first"]} {rows[0]["last"]} ({nid}): ردیف اصلی بین ردیف‌های تکراری مشخص نیست')
        done=self.log.completed_keys() if self.skip_done.isChecked() else set()
        self.queue=[e for e in auto_rows if (e['nid'],e['sheet'],str(e['row'])) not in done]
        if not self.queue: QMessageBox.information(self,"بدون عملیات","تمام ردیف‌های معتبر قبلاً با موفقیت ثبت شده‌اند."); return
        sample=self.queue[0]
        msg=(f"تعداد ردیف‌های آماده ثبت خودکار: {len(self.queue)}\nردیف‌های دومِ پرسنل دوشیفته که برای ثبت دستی کنار گذاشته شده‌اند: {len(manual_secondary)}\nردیف‌های دارای خطا یا ابهام که ثبت نمی‌شوند: {len(skipped)}\n\nاولین ردیف برای اجرای آزمایشی:\n{sample['first']} {sample['last']}\nکد ملی: {sample['nid']}\nنوع شیفت اکسل: {sample['shift']}\n\nثبت واقعی در سامانه انجام خواهد شد. ادامه می‌دهید؟")
        if QMessageBox.question(self,"تأیید شروع عملیات",msg,QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No)!=QMessageBox.StandardButton.Yes:return
        self.busy=True; self.index=0; self.stop_requested=False; self.start_btn.setEnabled(False); self.stop_btn.setEnabled(True); self.progress.setMaximum(len(self.queue)); self.progress.setValue(0); self.process_next()

    def request_stop(self):
        self.stop_requested=True; self.status.setText("درخواست توقف ثبت شد؛ پس از تعیین تکلیف ردیف جاری متوقف می‌شود.")
    def process_next(self):
        if not self.busy:return
        if self.stop_requested or self.index>=len(self.queue): self.finish(); return
        e=self.queue[self.index]; self.current=e; days,_=build_employee_plan(e,self.mapping); self._fill_started=False; self._verify_started=False
        self.status.setText(f"در حال آماده‌سازی {self.index+1}/{len(self.queue)}: {e['first']} {e['last']} — {e['nid']}")
        # Set the date range and open Telerik RadWindow through its client API.
        # The image opener click alone is unreliable in some QtWebEngine builds.
        js=f'''(() => {{
          try {{
            const s=document.getElementById('txtStartDate_TextBoxJalaliDate');
            const e=document.getElementById('txtEndDate_TextBoxJalaliDate');
            const img=document.getElementById('imagePersonnel');
            if(!s || !e || !img) return JSON.stringify({{ok:false,error:'کنترل‌های صفحه پیدا نشد',found:{{start:!!s,end:!!e,opener:!!img}},url:location.href}});
            s.value={json.dumps(self.data['start_date'])};
            e.value={json.dumps(self.data['end_date'])};
            const si=document.getElementById('txtStartIndex'); if(si) si.value='1';
            for (const el of [s,e]) {{
              el.dispatchEvent(new Event('input',{{bubbles:true}}));
              el.dispatchEvent(new Event('change',{{bubbles:true}}));
              if(window.jQuery) window.jQuery(el).trigger('change');
            }}
            const rw=window.$find ? window.$find('WindowPersonnel') : null;
            if(rw && typeof rw.show==='function') {{
              try {{ rw.setUrl(new URL('../MultySelectNew.aspx',location.href).href); }} catch(e) {{}}
              rw.show();
              return JSON.stringify({{ok:true,method:'RadWindow.show',frame:!!(rw.get_contentFrame && rw.get_contentFrame())}});
            }}
            img.click();
            return JSON.stringify({{ok:true,method:'image.click',radWindowFound:!!rw}});
          }} catch(err) {{ return JSON.stringify({{ok:false,error:String(err),stack:err && err.stack ? String(err.stack).slice(0,500):''}}); }}
        }})()'''
        self.js(js,self._picker_opened)

    def _picker_opened(self,result):
        if isinstance(result,str):
            try: result=json.loads(result)
            except Exception: result={'ok':False,'error':result}
        if not result or not result.get('ok'):
            detail=(result or {}).get('error','پاسخ مرورگر برای باز کردن پنجره خالی بود')
            extra=(result or {}).get('found')
            if extra: detail += ' — '+json.dumps(extra,ensure_ascii=False)
            self.fail_current('باز کردن پنجره انتخاب پرسنل: '+str(detail)); return
        self.status.setText('پنجره انتخاب پرسنل در حال باز شدن است...')
        QTimer.singleShot(1800,self.search_in_picker)
    def search_in_picker(self):
        nid=self.current['nid']
        js=f'''(() => {{
          const rw=window.$find && window.$find('WindowPersonnel');
          let doc=null;
          try {{ const fr=rw && rw.get_contentFrame(); doc=fr && fr.contentDocument; }} catch(e){{}}
          if(!doc) {{ const f=[...document.querySelectorAll('iframe')].find(x=>/MultySelectNew/i.test(x.src)); try{{doc=f&&f.contentDocument}}catch(e){{}} }}
          if(!doc) return {{ok:false,manual:true,error:'محتوای پنجره انتخاب پرسنل در دسترس نیست'}};
          const field=doc.getElementById('TextBoxCardNo');
          const btn=doc.getElementById('ButtonSearch');
          if(!field || !btn) return {{ok:false,manual:true,error:'فیلد شماره کارت (TextBoxCardNo) یا دکمه جستجو در HTML پنجره پیدا نشد'}};
          field.value={json.dumps(nid)};
          field.dispatchEvent(new Event('input',{{bubbles:true}}));
          field.dispatchEvent(new Event('change',{{bubbles:true}}));
          btn.click();
          return {{ok:true,searchField:field.id,searchButton:btn.id}};
        }})()'''
        self.js(js,self._searched)
    def _searched(self,res):
        if isinstance(res,str):
            try: res=json.loads(res)
            except Exception: res={'ok':False,'error':res}
        if not res or not res.get('ok'):
            self.ask_manual_picker((res or {}).get('error','جستجوی پرسنل خودکار انجام نشد')); return
        self._picker_wait_count=0
        QTimer.singleShot(500,self._wait_picker_results)
    def _wait_picker_results(self):
        if not self.busy or self.awaiting_manual_confirm:return
        self._picker_wait_count += 1
        nid=self.current['nid']
        js=f'''(() => {{
          const rw=window.$find&&window.$find('WindowPersonnel'); let d=null;
          try{{d=rw&&rw.get_contentFrame().contentDocument}}catch(e){{}}
          if(!d){{const f=[...document.querySelectorAll('iframe')].find(x=>/MultySelectNew/i.test(x.src));try{{d=f&&f.contentDocument}}catch(e){{}}}}
          if(!d)return {{ready:false,error:'iframe در دسترس نیست'}};
          const grid=d.getElementById('GridView1');
          const rows=grid?[...grid.querySelectorAll('tr')].filter(r=>(r.innerText||'').includes({json.dumps(nid)})):[];
          return {{ready:!!grid,rows:rows.length,grid:!!grid}};
        }})()'''
        self.js(js,lambda r:self._picker_results_checked(r))
    def _picker_results_checked(self,res):
        if isinstance(res,str):
            try: res=json.loads(res)
            except Exception: res={}
        if res and res.get('ready'):
            self.select_picker_result(); return
        if getattr(self,'_picker_wait_count',0) < 40:
            QTimer.singleShot(500,self._wait_picker_results)
        else:
            self.ask_manual_picker('نتایج جستجوی پرسنل پس از ۲۰ ثانیه بارگذاری نشد؛ لطفاً نتیجه را در پنجره بررسی کنید.')
    def select_picker_result(self):
        nid=self.current['nid']
        js=f'''(() => {{
          const rw=window.$find&&window.$find('WindowPersonnel'); let d=null;
          try{{d=rw&&rw.get_contentFrame().contentDocument}}catch(e){{}}
          if(!d){{const f=[...document.querySelectorAll('iframe')].find(x=>/MultySelectNew/i.test(x.src));try{{d=f&&f.contentDocument}}catch(e){{}}}}
          if(!d)return {{ok:false,error:'iframe انتخاب پرسنل در دسترس نیست'}};
          const grid=d.getElementById('GridView1');
          if(!grid)return {{ok:false,error:'جدول نتایج GridView1 پیدا نشد'}};
          const rows=[...grid.querySelectorAll('tr')].filter(r=>{{
            const cells=[...r.querySelectorAll('td')].map(c=>(c.innerText||'').trim());
            return cells.some(t=>t==={json.dumps(nid)});
          }});
          if(rows.length!==1)return {{ok:false,error:rows.length===0?'نتیجه‌ای با کد ملی دقیق پیدا نشد':'بیش از یک ردیف با کد ملی یکسان پیدا شد؛ انتخاب خودکار لغو شد',matches:rows.length}};
          const row=rows[0]; const cb=row.querySelector('input[type=checkbox]');
          if(!cb)return {{ok:false,error:'چک‌باکس انتخاب ردیف پیدا نشد'}};
          if(!cb.checked)cb.click();
          const hf=d.getElementById('hf1');
          if(!hf || !hf.value || hf.value===',')return {{ok:false,error:'انتخاب در فیلد پنهان hf1 تأیید نشد'}};
          const count=d.getElementById('Counts');
          if(count && Number(count.value)!==1)return {{ok:false,error:'تعداد افراد انتخاب‌شده برابر یک نیست',count:count.value}};
          const confirm=d.getElementById('ButtonOk')||d.getElementById('ButtonOKTop');
          if(!confirm)return {{ok:false,error:'دکمه تأیید واقعی ButtonOk پیدا نشد'}};
          confirm.click();
          return {{ok:true,confirmed:true,selectedCount:count?count.value:'1',confirmId:confirm.id}};
        }})()'''
        self.js(js,self._selected)
    def _selected(self,res):
        if isinstance(res,str):
            try: res=json.loads(res)
            except Exception: res={'ok':False,'error':res}
        if not res or not res.get('ok'):
            self.ask_manual_picker((res or {}).get('error','انتخاب ردیف پرسنل انجام نشد')); return
        if not res.get('confirmed'):
            self.ask_manual_picker('پرسنل انتخاب شد؛ دکمه سبز تأیید در پنجره را دستی بزنید.'); return
        QTimer.singleShot(1100,self.prepare_days)
    def ask_manual_picker(self,reason):
        self.awaiting_manual_confirm=True; self._manual_stage='picker'
        if hasattr(self,'manual_btn'):
            self.manual_btn.setParent(None); self.manual_btn.deleteLater()
        self.status.setText(reason+' پس از تأیید پنجره، دکمه «ادامه پس از انتخاب پرسنل» را بزنید.')
        self.manual_btn=QPushButton("ادامه پس از انتخاب پرسنل")
        self.manual_btn.clicked.connect(self.continue_after_manual)
        self.status.parentWidget().layout().addWidget(self.manual_btn)
        self.browser.setFocus()
    def continue_after_manual(self):
        if hasattr(self,'manual_btn'): self.manual_btn.setParent(None); self.manual_btn.deleteLater()
        self.awaiting_manual_confirm=False
        if getattr(self,'_manual_stage','')=='show_days':
            self.begin_wait_for_days()
        else:
            QTimer.singleShot(700,self.prepare_days)
    def ask_manual_show_days(self,reason):
        self.awaiting_manual_confirm=True; self._manual_stage='show_days'
        if hasattr(self,'manual_btn'):
            self.manual_btn.setParent(None); self.manual_btn.deleteLater()
        self.status.setText(reason+' در مرورگر سامانه دکمه «نمایش روزهای شیفت» را دستی بزنید؛ سپس دکمه ادامه را در همین برنامه انتخاب کنید.')
        self.manual_btn=QPushButton('ادامه پس از نمایش روزهای شیفت')
        self.manual_btn.clicked.connect(self.continue_after_manual)
        self.status.parentWidget().layout().addWidget(self.manual_btn)
        self.browser.setFocus()
    def prepare_days(self):
        # Check for an already-rendered day grid. Otherwise activate the page's
        # show-days control by a known ID or its visible Persian label.
        self.status.setText('بررسی دکمه نمایش روزهای شیفت و آماده‌سازی تقویم...')
        js=r'''(() => {
          const norm=s=>(s||'').replace(/\s+/g,' ').trim();
          const realDayFields=[...document.querySelectorAll('input[id^="rpMonths_ctl00_rpDays_ctl"][id$="_txtShiftName"]')].filter(x=>{const m=x.id.match(/^rpMonths_ctl00_rpDays_ctl(\d{2})_txtShiftName$/);return !!m && Number(m[1])>=6 && Number(m[1])<=35;});
          const realDayContainers=[...Array.from({length:30},(_,i)=>document.getElementById('rpMonths_ctl00_rpDays_ctl'+String(i+6).padStart(2,'0')+'_divDayITem'))].filter(Boolean);
          if(realDayFields.length===30 && realDayContainers.length===30) return {ok:true,alreadyVisible:true,fields:realDayFields.length,containers:realDayContainers.length,url:location.href};
          const ids=['btShowFinalDaysByDate','btShowDays','btnShowDays','btnShowFinalDaysByDate'];
          let target=ids.map(id=>document.getElementById(id)).find(Boolean)||null;
          if(!target) {
            const controls=[...document.querySelectorAll('input[type=submit],input[type=button],button,a,[role=button]')];
            const matches=controls.filter(el=>/نمایش\s*روزهای?\s*شیفت|روزهای?\s*شیفت\s*را\s*نمایش/.test(norm(el.innerText||el.value||el.getAttribute('aria-label'))));
            if(matches.length===1) target=matches[0];
            else return {ok:false,alreadyVisible:false,dates:0,candidates:matches.map(x=>({id:x.id,text:norm(x.innerText||x.value),tag:x.tagName})),url:location.href};
          }
          const info={id:target.id||'',text:norm(target.innerText||target.value),tag:target.tagName};
          target.click();
          return {ok:true,clicked:true,control:info,url:location.href};
        })()'''
        self.js(js,self._show_days_triggered)
    def _show_days_triggered(self,res):
        if isinstance(res,str):
            try: res=json.loads(res)
            except Exception: res={'ok':False,'error':res}
        if res and res.get('ok'):
            if res.get('alreadyVisible'):
                self.status.setText(f"۳۰ کادر روزها از قبل در صفحه وجود دارد. در حال پرکردن ctl06 تا ctl35...")
            else:
                self.status.setText('دکمه نمایش روزهای شیفت اجرا شد؛ در انتظار بارگذاری تقویم...')
            QTimer.singleShot(700,self.begin_wait_for_days)
            return
        detail='دکمه نمایش روزهای شیفت در صفحه پیدا نشد.'
        if res and res.get('candidates'):
            detail+=' کنترل‌های مشابه: '+json.dumps(res['candidates'],ensure_ascii=False)
        elif res and res.get('error'):
            detail+=' '+str(res['error'])
        self.ask_manual_show_days(detail)
    def begin_wait_for_days(self):
        self._show_wait_count=0
        self.status.setText('در انتظار آماده‌شدن ۳۰ کادر روزهای شیفت (ctl06 تا ctl35)...')
        QTimer.singleShot(250,self._wait_for_days)
    def show_days(self):
        self.begin_wait_for_days()
    def _wait_for_days(self):
        if self._fill_started or not self.busy:return
        self._show_wait_count += 1
        js=r'''(() => {
          const fields=[...document.querySelectorAll('input[id^="rpMonths_ctl00_rpDays_ctl"][id$="_txtShiftName"]')].filter(f=>{
            const m=f.id.match(/^rpMonths_ctl00_rpDays_ctl(\d{2})_txtShiftName$/);
            return !!m && Number(m[1])>=6 && Number(m[1])<=35;
          });
          const paired=fields.filter(f=>{const m=f.id.match(/ctl(\d{2})_txtShiftName$/);return !!m && !!document.getElementById('rpMonths_ctl00_rpDays_ctl'+m[1]+'_divDayITem');});
          return {fields:fields.length,paired:paired.length};
        })()'''
        self.js(js,self._days_presence_checked)
    def _days_presence_checked(self,res):
        if isinstance(res,str):
            try: res=json.loads(res)
            except Exception: res={}
        if res and res.get('fields',0)==30 and res.get('paired',0)==30:
            QTimer.singleShot(250,self._fill_days_once)
        elif self._show_wait_count < 60:
            QTimer.singleShot(500,self._wait_for_days)
        else:
            self.fail_current('پس از ۳۰ ثانیه، ۳۰ کادر واقعی روزهای شیفت (ctl06 تا ctl35) در صفحه پیدا نشدند. احتمالاً تقویم روزها بارگذاری نشده است.')
    def _fill_days_once(self):
        if self._fill_started:return
        self._fill_started=True; self._filling=True
        days,plan_errors=build_employee_plan(self.current,self.mapping)
        if plan_errors:
            self.fail_current('اعتبارسنجی برنامه شیفت ناموفق بود: '+'؛ '.join(plan_errors)); return
        if len(days)!=30 or set(days.keys())!={str(i) for i in range(1,31)}:
            self.fail_current(f'تعداد روزهای برنامه اکسل باید دقیقاً ۳۰ روز باشد؛ تعداد فعلی: {len(days)}.'); return
        js=fr'''(() => {{
          const map={json.dumps({str(k):v for k,v in days.items()},ensure_ascii=False)};
          const targetYear={json.dumps(str(self.data['year']))};
          const targetMonth={json.dumps(f'{self.data["month_number"]:02d}')};
          const fields=[...document.querySelectorAll('input[id^="rpMonths_ctl00_rpDays_ctl"][id$="_txtShiftName"]')].filter(f=>{{
            const m=f.id.match(/^rpMonths_ctl00_rpDays_ctl(\d{{2}})_txtShiftName$/);
            return !!m && Number(m[1])>=6 && Number(m[1])<=35;
          }}).sort((a,b)=>Number(a.id.match(/ctl(\d{{2}})_txtShiftName$/)[1])-Number(b.id.match(/ctl(\d{{2}})_txtShiftName$/)[1]));
          if(fields.length!==30)return {{ok:false,error:'تعداد کادرهای واقعی روزها ۳۰ نیست',found:fields.length,expected:30}};
          const paired=[]; const missingDays=[]; const mismatches=[];
          for(let i=0;i<30;i++){{
            const index=i+6; const day=String(i+1); const field=fields[i];
            const expectedId='rpMonths_ctl00_rpDays_ctl'+String(index).padStart(2,'0')+'_txtShiftName';
            const container=document.getElementById('rpMonths_ctl00_rpDays_ctl'+String(index).padStart(2,'0')+'_divDayITem');
            if(field.id!==expectedId || !container){{missingDays.push(day);continue;}}
            const dateField=document.getElementById('rpMonths_ctl00_rpDays_ctl'+String(index).padStart(2,'0')+'_hfFaDate');
            const date=(dateField?.value||'').replace(/[۰-۹٠-٩]/g,c=>String('۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩'.indexOf(c)%10));
            // If the site supplies a valid hidden date, use it as a safety check only.
            // The field-to-day mapping itself is deterministic: ctl06=day 1 ... ctl35=day 30.
            if(/^14\d{{2}}\/\d{{2}}\/\d{{2}}$/.test(date)){{
              const parts=date.split('/');
              if(parts[0]!==targetYear || parts[1]!==targetMonth || String(parseInt(parts[2],10))!==day){{mismatches.push({{index,day,date}});continue;}}
            }}
            if(!Object.prototype.hasOwnProperty.call(map,day)){{missingDays.push(day);continue;}}
            paired.push({{field,day}});
          }}
          if(missingDays.length||mismatches.length)return {{ok:false,error:'تطبیق کادرهای روزها ناقص یا تاریخ صفحه با ماه اکسل متفاوت است',missingDays,mismatches,found:paired.length,expected:30}};
          for(const item of paired){{
            const setter=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value').set;
            setter.call(item.field,map[item.day]);
            item.field.dispatchEvent(new Event('input',{{bubbles:true}}));
            item.field.dispatchEvent(new Event('change',{{bubbles:true}}));
            item.field.dispatchEvent(new Event('blur',{{bubbles:true}}));
          }}
          const valueMismatches=paired.filter(x=>x.field.value!==map[x.day]).map(x=>x.day);
          return {{ok:valueMismatches.length===0,filled:paired.length,expected:30,planned:Object.keys(map).length,unmatched:missingDays.length,mismatches:valueMismatches,detector:'fixed-index-ctl06-ctl35',matchedDays:paired.map(x=>x.day)}};
        }})()'''
        self.js(js,self._days_filled)
    def _days_filled(self,res):
        self._filling=False
        if isinstance(res,str):
            try: res=json.loads(res)
            except Exception: res={'ok':False,'error':res}
        if not res or not res.get('ok'):
            self.fail_current((res or {}).get('error','ورود روزها ناموفق بود')+' '+json.dumps({k:v for k,v in (res or {}).items() if k!='error'},ensure_ascii=False)); return
        if res.get('filled',0)<1: self.fail_current('هیچ روزی با تاریخ اکسل تطبیق داده نشد.'); return
        if res.get('filled') != res.get('expected') or res.get('expected') != res.get('planned') or res.get('unmatched',0)>0 or res.get('mismatches'):
            self.fail_current(f"تطبیق روزها ناقص است (پرشده: {res.get('filled')}، مورد انتظار: {res.get('expected')}، بدون جفت تاریخ: {res.get('unmatched')}، اختلاف مقادیر: {res.get('mismatches')}). ارسال انجام نشد."); return
        self.status.setText(f"{self.current['first']} {self.current['last']}: {res['filled']} روز از {res.get('total',0)} کادر بررسی و پر شد؛ در حال ثبت در سامانه...")
        try: self.browser.loadFinished.disconnect(self._after_submit_load)
        except Exception: pass
        self._verify_attempts=0
        self.browser.loadFinished.connect(self._after_submit_load)
        self.js("(() => { const b=document.getElementById('btSetShiftsDay')||document.getElementById('btSetShiftsDay2'); if(!b)return {ok:false,error:'دکمه ثبت نهایی پیدا نشد'}; const baselineHasSuccess=(document.body.innerText||'').includes('روز های الگو با موفقیت ذخیره شدند.'); b.click(); return {ok:true,baselineHasSuccess}; })()",self._submit_clicked)
    def _submit_clicked(self,res):
        if isinstance(res,str):
            try: res=json.loads(res)
            except Exception: res={'ok':False,'error':res}
        if not res or not res.get('ok'):
            try: self.browser.loadFinished.disconnect(self._after_submit_load)
            except Exception: pass
            self.fail_current((res or {}).get('error','ثبت نهایی انجام نشد')); return
        self._submit_baseline_has_success=bool(res.get('baselineHasSuccess',False))
        QTimer.singleShot(1200,self._verify_submit)
    def _after_submit_load(self,ok):
        try:self.browser.loadFinished.disconnect(self._after_submit_load)
        except Exception:pass
        QTimer.singleShot(500,self._verify_submit)
    def _verify_submit(self):
        if not self.busy:return
        self._verify_attempts=getattr(self,'_verify_attempts',0)+1
        js="(() => { const t=(document.body.innerText||''); const ok=t.includes('روز های الگو با موفقیت ذخیره شدند.'); const bad=t.includes('خطا')||t.includes('ناموفق'); return {ok,bad,text:t.slice(-1800)}; })()"
        self.js(js,self._verified)
    def _verified(self,res):
        if isinstance(res,str):
            try: res=json.loads(res)
            except Exception: res={'ok':False,'text':res}
        if res and res.get('ok') and not res.get('bad') and not getattr(self,'_submit_baseline_has_success',False):
            self.log.append(self.current,'موفق','پیام موفقیت سامانه مشاهده شد')
            self.status.setText(f"ثبت موفق: {self.current['first']} {self.current['last']}")
            self.index+=1; self.progress.setValue(self.index); self._filling=False; QTimer.singleShot(800,self.process_next)
        elif getattr(self,'_verify_attempts',0) < 20:
            self.status.setText(f"در انتظار تأیید نهایی سامانه... ({self._verify_attempts}/20)")
            QTimer.singleShot(700,self._verify_submit)
        else:
            detail=(res or {}).get('text','پیام موفقیت قطعی مشاهده نشد')
            if getattr(self,'_submit_baseline_has_success',False): detail='پیام موفقیت پیش از ثبت نیز در صفحه وجود داشت؛ نتیجه جدید قابل اثبات نیست. '+detail
            self.log.append(self.current,'نیازمند بررسی',detail[-500:])
            self.status.setText('پیام موفقیت قطعی مشاهده نشد؛ برای جلوگیری از ثبت تکراری عملیات متوقف شد.')
            QMessageBox.warning(self,'تأیید ثبت نامشخص','سامانه پیام موفقیت قطعی نشان نداد. برای جلوگیری از ثبت تکراری، ادامه متوقف شد. وضعیت این ردیف در گزارش با «نیازمند بررسی» ثبت شد.')
            self.index+=1; self.progress.setValue(self.index); self.finish()
    def fail_current(self,detail):
        self.log.append(self.current,'خطا',detail); self.status.setText('خطا: '+detail); QMessageBox.warning(self,'ثبت متوقف شد',detail+'\nبرای جلوگیری از ثبت ناقص یا تکراری، عملیات متوقف شد.'); self.finish()
    def finish(self):
        self.busy=False; self.start_btn.setEnabled(True); self.stop_btn.setEnabled(False)
        self.status.setText(f"عملیات پایان یافت. {self.index} ردیف از {len(self.queue)} پردازش شد. گزارش: {self.log.path}")
        QMessageBox.information(self,'پایان عملیات',self.status.text())
    def inspect_page(self):
        js="(() => ({url:location.href,title:document.title,controls:[...document.querySelectorAll('input,button,select,iframe')].map(x=>({id:x.id,name:x.name,type:x.type,value:x.type==='password'?'':(x.value||'').slice(0,50),src:x.src||''})),text:(document.body.innerText||'').slice(0,2500)}))()"
        self.js(js,lambda r: QMessageBox.information(self,'وضعیت صفحه',json.dumps(r,ensure_ascii=False,indent=2)[:6000]))

if __name__=="__main__":
    app=QApplication(sys.argv); app.setLayoutDirection(Qt.LayoutDirection.RightToLeft); win=MainWindow(); win.show(); sys.exit(app.exec())
