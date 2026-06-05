#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
UDF Wizard  -  Offline masaüstü uygulaması (modern arayüz)
Word (.docx) ve PDF (.pdf) <-> UYAP UDF (.udf) çift yönlü dönüştürme.

Çalıştırmak için:  python udf_donusturucu.py
Gereksinimler   :  pip install customtkinter python-docx pdfplumber reportlab
(tkinter Python ile birlikte gelir; ek kurulum gerekmez.)
"""

import os
import sys
import threading

import customtkinter as ctk
from tkinter import filedialog, messagebox

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import udf_core as core


SUPPORTED_IN = [("Belgeler", "*.udf *.docx *.pdf"),
                ("UDF", "*.udf"), ("Word", "*.docx"), ("PDF", "*.pdf")]

TARGETS = {
    ".udf":  [("Word (.docx)", ".docx"), ("PDF (.pdf)", ".pdf")],
    ".docx": [("UDF (.udf)", ".udf"),    ("PDF (.pdf)", ".pdf")],
    ".pdf":  [("UDF (.udf)", ".udf"),    ("Word (.docx)", ".docx")],
}

TYPE_INFO = {
    ".udf":  ("UDF",  "#7048e8"),
    ".docx": ("Word", "#1971c2"),
    ".pdf":  ("PDF",  "#e03131"),
}

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("UDF Wizard")
        self.geometry("680x540")
        self.minsize(620, 500)

        self.in_paths = []
        self._targets = []
        self._dark = False

        self.grid_columnconfigure(0, weight=1)
        self._build()

    def _build(self):
        header = ctk.CTkFrame(self, corner_radius=0, fg_color=("#3b5bdb", "#2b3a67"), height=92)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)
        header.grid_propagate(False)

        ctk.CTkLabel(header, text="UDF Wizard",
                     font=ctk.CTkFont(size=24, weight="bold"),
                     text_color="white").grid(row=0, column=0, sticky="w", padx=24, pady=(18, 0))
        ctk.CTkLabel(header, text="Word / PDF  <->  UYAP UDF   •   Tamamen çevrimdışı",
                     font=ctk.CTkFont(size=13), text_color="#dfe3ff").grid(
                         row=1, column=0, sticky="w", padx=24, pady=(0, 16))

        self.theme_btn = ctk.CTkButton(header, text="Koyu Tema", width=96, height=32,
                                       text_color="#ffffff", fg_color="#555555", hover_color="#444444",
                                       command=self._toggle_theme)
        self.theme_btn.grid(row=0, column=1, rowspan=2, padx=20)

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", padx=24, pady=20)
        body.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        card1 = ctk.CTkFrame(body, corner_radius=14)
        card1.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        card1.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(card1, text="1  •  Dönüştürülecek dosya",
                     font=ctk.CTkFont(size=14, weight="bold")).grid(
                         row=0, column=0, columnspan=2, sticky="w", padx=18, pady=(14, 4))
        self.lbl_in = ctk.CTkLabel(card1, text="Henüz dosya seçilmedi.",
                                   font=ctk.CTkFont(size=13), anchor="w",
                                   text_color=("#666", "#aaa"))
        self.lbl_in.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 16))
        ctk.CTkButton(card1, text="Dosya Seç  (çoklu)", width=150, height=40,
                      font=ctk.CTkFont(size=13, weight="bold"),
                      corner_radius=10, command=self.pick_input).grid(
                          row=1, column=1, padx=18, pady=(0, 16))

        card2 = ctk.CTkFrame(body, corner_radius=14)
        card2.grid(row=1, column=0, sticky="ew", pady=(0, 16))
        card2.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(card2, text="2  •  Hedef format",
                     font=ctk.CTkFont(size=14, weight="bold")).grid(
                         row=0, column=0, sticky="w", padx=18, pady=(14, 4))
        self.seg = ctk.CTkSegmentedButton(card2, values=["—"],
                                          font=ctk.CTkFont(size=13),
                                          height=40, corner_radius=10)
        self.seg.grid(row=1, column=0, sticky="w", padx=18, pady=(0, 16))
        self.seg.configure(state="disabled")

        self.btn_go = ctk.CTkButton(body, text="Dönüştür", height=50,
                                    font=ctk.CTkFont(size=16, weight="bold"),
                                    corner_radius=12, fg_color="#2f9e44",
                                    hover_color="#2b8a3e", command=self.run_convert,
                                    state="disabled")
        self.btn_go.grid(row=2, column=0, sticky="ew", pady=(4, 4))

        self.progress = ctk.CTkProgressBar(body, height=6, corner_radius=3)
        self.progress.grid(row=3, column=0, sticky="ew", pady=(12, 4))
        self.progress.set(0)
        self.progress.grid_remove()

        self.status = ctk.CTkLabel(body, text="", font=ctk.CTkFont(size=13),
                                   text_color=("#444", "#bbb"))
        self.status.grid(row=4, column=0, sticky="w", pady=(2, 0))

        ctk.CTkLabel(self,
                     text="PDF -> UDF dönüşümünde metin korunur; karmaşık görsel düzen birebir aktarılmayabilir.",
                     font=ctk.CTkFont(size=11), text_color=("#999", "#777"),
                     wraplength=620, justify="left").grid(
                         row=2, column=0, sticky="w", padx=24, pady=(0, 14))

    def _toggle_theme(self):
        self._dark = not self._dark
        ctk.set_appearance_mode("dark" if self._dark else "light")
        if self._dark:
            self.theme_btn.configure(text="Açık Tema", text_color="#000000",
                                     fg_color="#cccccc", hover_color="#bbbbbb")
        else:
            self.theme_btn.configure(text="Koyu Tema", text_color="#ffffff",
                                     fg_color="#555555", hover_color="#444444")

    def pick_input(self):
        paths = filedialog.askopenfilenames(title="Dönüştürülecek dosya(lar)ı seçin",
                                            filetypes=SUPPORTED_IN)
        if not paths:
            return
        exts = {os.path.splitext(p)[1].lower() for p in paths}
        if len(exts) > 1:
            messagebox.showerror("Hata", "Toplu dönüştürmede tüm dosyalar aynı türde olmalı "
                                 "(hepsi .docx veya hepsi .pdf gibi).")
            return
        ext = exts.pop()
        if ext not in TARGETS:
            messagebox.showerror("Hata", "Desteklenmeyen dosya türü: %s\nYalnızca .udf, .docx, .pdf." % ext)
            return

        self.in_paths = list(paths)
        type_label, _ = TYPE_INFO.get(ext, (ext, "#666"))
        if len(paths) == 1:
            self.lbl_in.configure(text="%s   (%s)" % (os.path.basename(paths[0]), type_label),
                                  text_color=("#222", "#eee"))
        else:
            self.lbl_in.configure(text="%d dosya seçildi   (%s)" % (len(paths), type_label),
                                  text_color=("#222", "#eee"))

        opts = TARGETS[ext]
        self._targets = opts
        values = [o[0] for o in opts]
        self.seg.configure(values=values, state="normal")
        self.seg.set(values[0])
        self.btn_go.configure(state="normal")
        self.status.configure(text="Hazır. Hedef formatı seçip Dönüştür'e basın.",
                              text_color=("#444", "#bbb"))

    def run_convert(self):
        if not self.in_paths or not self._targets:
            return
        sel = self.seg.get()
        target = next((o for o in self._targets if o[0] == sel), self._targets[0])
        target_label, target_ext = target

        if len(self.in_paths) == 1:
            base = os.path.splitext(os.path.basename(self.in_paths[0]))[0]
            out_path = filedialog.asksaveasfilename(
                title="Çıktı dosyasını kaydet", defaultextension=target_ext,
                initialfile=base + target_ext,
                filetypes=[(target_label, "*" + target_ext)])
            if not out_path:
                return
            jobs = [(self.in_paths[0], out_path)]
        else:
            out_dir = filedialog.askdirectory(title="Çıktı klasörünü seçin")
            if not out_dir:
                return
            jobs = [(ip, os.path.join(out_dir,
                     os.path.splitext(os.path.basename(ip))[0] + target_ext))
                    for ip in self.in_paths]

        self.btn_go.configure(state="disabled")
        self.progress.grid()
        self.progress.set(0)
        self.progress.start()
        self.status.configure(text="Dönüştürülüyor…", text_color=("#3b5bdb", "#8aa0ff"))
        threading.Thread(target=self._worker, args=(jobs,), daemon=True).start()

    def _worker(self, jobs):
        ok, errors = 0, []
        for ip, op in jobs:
            try:
                core.convert(ip, op)
                ok += 1
            except Exception as e:
                errors.append((os.path.basename(ip), str(e)))
        self.after(0, lambda: self._done(ok, errors))

    def _done(self, ok, errors):
        self.progress.stop()
        self.progress.set(1)
        self.progress.grid_remove()
        self.btn_go.configure(state="normal")
        if not errors:
            self.status.configure(text="Tamamlandı  (%d dosya)" % ok,
                                  text_color=("#2f9e44", "#69db7c"))
            messagebox.showinfo("Başarılı", "%d dosya başarıyla dönüştürüldü." % ok)
        else:
            self.status.configure(text="%d başarılı   •   %d hatalı" % (ok, len(errors)),
                                  text_color=("#e8590c", "#ffa94d"))
            detail = "\n".join("• %s: %s" % (n, m) for n, m in errors[:8])
            extra = ""
            if any("yüklü değil" in m for _, m in errors):
                extra = "\n\nGerekli kütüphaneyi kurun:\n  pip install customtkinter python-docx pdfplumber reportlab"
            messagebox.showwarning("Kısmen tamamlandı",
                                   "%d dosya dönüştürüldü.\n\nHatalı dosyalar:\n%s%s" % (ok, detail, extra))


if __name__ == "__main__":
    App().mainloop()
