import os
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.core.window import Window
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

Window.clearcolor = (0.94, 0.97, 0.96, 1)
PROTOCOL_PASSWORD = "111303"
COMPANY_NAME = "GUANGZHOU HAIJINSHUN TRADE CO.,LTD."

class AhmedProtocolUI(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.padding = 20
        self.spacing = 15

        self.title_lbl = Label(
            text="نظام المعالجة الجمركية\n(AHMED PROTOCOL)",
            font_size='18sp', bold=True, color=(0.18, 0.35, 0.15, 1), halign='center'
        )
        self.add_widget(self.title_lbl)

        self.btn_inv = Button(text="[+] الفاتورة التجارية (Invoice)", size_hint_y=None, height=50)
        self.btn_inv.bind(on_press=lambda x: self.select_doc(self.btn_inv, "الفاتورة"))
        self.add_widget(self.btn_inv)

        self.btn_co = Button(text="[+] شهادة المنشأ (CO)", size_hint_y=None, height=50)
        self.btn_co.bind(on_press=lambda x: self.select_doc(self.btn_co, "شهادة المنشأ"))
        self.add_widget(self.btn_co)

        self.btn_pl = Button(text="[+] بيان التعبئة (Packing List)", size_hint_y=None, height=50)
        self.btn_pl.bind(on_press=lambda x: self.select_doc(self.btn_pl, "بيان التعبئة"))
        self.add_widget(self.btn_pl)

        self.status_lbl = Label(text="بانتظار تجهيز المستندات...", font_size='13sp', color=(0.4, 0.4, 0.4, 1))
        self.add_widget(self.status_lbl)

        self.btn_run = Button(text="⚡ توليد وتشفير البيان الجمركي", size_hint_y=None, height=60, bold=True)
        self.btn_run.bind(on_press=self.generate_excel)
        self.add_widget(self.btn_run)

    def select_doc(self, btn, name):
        btn.text = f"✅ تم تحديد {name}"

    def show_msg(self, title, msg):
        box = BoxLayout(orientation='vertical', padding=10, spacing=10)
        box.add_widget(Label(text=msg, halign='center'))
        b = Button(text="موافق", size_hint_y=None, height=40)
        box.add_widget(b)
        p = Popup(title=title, content=box, size_hint=(0.85, 0.4))
        b.bind(on_press=p.dismiss)
        p.open()

    def generate_excel(self, instance):
        try:
            try:
                from android.storage import primary_external_storage_path
                base_dir = os.path.join(primary_external_storage_path(), "Download")
            except Exception:
                base_dir = os.getcwd()

            file_path = os.path.join(base_dir, "FO20261005.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "البيان الجمركي"
            ws.views.sheetView.rightToLeft = True

            ws.merge_cells("A1:G1")
            head = ws["A1"]
            head.value = COMPANY_NAME
            head.font = Font(name="Arial", size=13, bold=True, color="FFFFFF")
            head.fill = PatternFill(start_color="3B7A57", end_color="3B7A57", fill_type="solid")
            head.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 35

            headers = ["الرمز المنسق", "عدد الطرود", "الوصف", "القيمة بالدولار", "عدد الحبات", "الوزن الصافي", "الوزن القائم"]
            for col_idx, h in enumerate(headers, 1):
                c = ws.cell(row=2, column=col_idx, value=h)
                c.fill = PatternFill(start_color="A3C1AD", end_color="A3C1AD", fill_type="solid")
                c.alignment = Alignment(horizontal="center", vertical="center")

            dummy = [
                ["6109.10", 50, "تي شيرت قطن رجالي", 5000, 1000, 250, 270],
                ["6403.99", 30, "أحذية جلد كلاسيك", 4500, 600, 180, 195]
            ]
            r_idx = 3
            for r in dummy:
                for c_idx, val in enumerate(r, 1):
                    ws.cell(row=r_idx, column=c_idx, value=val)
                r_idx += 1

            ws.cell(row=r_idx, column=3, value="الإجمالي (SUM)").font = Font(bold=True)
            for c_num in [2, 4, 5, 6, 7]:
                col_let = get_column_letter(c_num)
                ws.cell(row=r_idx, column=c_num, value=f"=SUM({col_let}3:{col_let}{r_idx-1})")

            ws.protection.set_password(PROTOCOL_PASSWORD)
            wb.security.workbookPassword = PROTOCOL_PASSWORD
            wb.save(file_path)

            self.status_lbl.text = "✅ تم إنشاء الملف بنجاح!"
            self.show_msg("نجاح", f"تم حفظ الملف في:\n{file_path}")
        except Exception as e:
            self.show_msg("خطأ", str(e))

class CustomsApp(App):
    def build(self):
        return AhmedProtocolUI()

if __name__ == '__main__':
    CustomsApp().run()
