import tkinter as tk
import xlsxwriter
import os
from openpyxl import load_workbook
from tkinter import *
from random import random
# from math import exp, log, sqrt
from itertools import permutations
# from tkinter.ttk import Progressbar
# import time
# import random

global msayisi
global ssayisi
global nehgantsira


def degergetir():
    global msayisi
    global ssayisi
    sonuc2 = Label(text=str(issayisi.get()) + '    /    ' + str(makinesayisi.get()) + '    /   F   /   Fmax   ',
                   fg='orange', bg='purple', width=40, font=196)
    sonuc2.place(x=720, y=350)
    msayisi = int(makinesayisi.get())
    ssayisi = int(issayisi.get())


def sonucac():
    os.system("sonuclarrr.xlsx")


def GANT(algoritma_adi, sonuc, sayfano, yazmasirasi_2):
    sayac11 = 0
    gliste1 = []
    gliste2 = []
    glistesonuc = []
    isindex = []
    sonliste = []
    toplam = 0
    sayacindex = 0
    bekleme = 0
    zaman = 0
    guncelleme = 0
    sayac9 = 0
    sayac10 = 0
    sayacgant = 3
    yazmasirasi = 0

    while sayac11 < is_sayisi:
        sayfano.write(yazmasirasi, sayac11, sonuc[sayac11], renklendirme2)
        sayac11 += 1
    sayac11 = 0
    yazmasirasi += 1

    for cell in sayfa['2']:
        if cell.value is not None:
            gliste1.append(cell.value)
    gliste1.pop(0)
    while sayacindex < is_sayisi:
        isindex.append(sonuc[sayacindex] - 1)
        toplam = toplam + gliste1[isindex[sayacindex]]
        glistesonuc.append(toplam)
        sayacindex += 1
        # ilk makinenin gant aşaması tamamlandı

    while sayac11 < is_sayisi:
        sayfano.write(yazmasirasi, sayac11, glistesonuc[sayac11])
        sayac11 += 1
    sayac11 = 0
    yazmasirasi += 1

    while sayacgant < makine_sayisi + 2:
        for cell in sayfa[sayacgant]:
            if cell.value is not None:
                gliste2.append(cell.value)
        gliste2.pop(0)
        while sayac9 < is_sayisi:
            while glistesonuc[sayac9] > bekleme:
                bekleme = bekleme + 1
            zaman = bekleme + gliste2[isindex[sayac9]]
            sonliste.append(zaman)
            guncelleme = sonliste[sayac9 - 1] + gliste2[isindex[sayac9]]
            if sonliste[sayac9] < guncelleme:
                sonliste[sayac9] = guncelleme
            if sayac9 == 0:
                sonliste[sayac9] = zaman
            sayac9 += 1
            zaman = 0
            bekleme = 0
        while sayac10 < len(sonliste):
            glistesonuc[sayac10] = sonliste[sayac10]
            sayac10 += 1
        while sayac11 < is_sayisi:
            sayfano.write(yazmasirasi, sayac11, glistesonuc[sayac11])
            sayac11 += 1
        yazmasirasi += 1
        sayac11 = 0
        sonliste.clear()
        # print(glistesonuc)
        sayac9 = 0
        sayac10 = 0
        gliste2.clear()
        sayacgant += 1
    sayacgant = 3
    guncelleme = 0
    sonuc_sira.clear()
    sonuc_sira.append(sonuc)
    sonuc_fmax.clear()
    sonuc_fmax.append(max(glistesonuc))
    sayfa6.write(yazmasirasi_2, 0, algoritma_adi, renklendirme)
    sayfa6.write(yazmasirasi_2, 1, sonuc_fmax[0], renklendirme1)
    sayfa6.write_row(yazmasirasi_2, 2, sonuc_sira[0])


def NEH(yazmasirasi):
    sayac11 = 0
    while sayac11 < len(Neh_gant):
        sayfa5.write(yazmasirasi, sayac11, Neh_gant[sayac11], renklendirme2)
        sayac11 += 1
    sayac11 = 0
    yazmasirasi += 1
    gliste1 = []
    gliste2 = []
    global nehgantsira
    nehgantsira = []
    isindex = []
    sonliste = []
    toplam = 0
    sayacindex = 0
    bekleme = 0
    zaman = 0
    guncelleme = 0
    sayac9 = 0
    sayac10 = 0
    sayacgant = 3
    bitis = makine_sayisi + 1

    # işleri index islemine düzeltme
    for cell in sayfa['2']:
        if cell.value is not None:
            gliste1.append(cell.value)
    gliste1.pop(0)
    while sayacindex < len(Neh_gant):
        isindex.append(Neh_gant[sayacindex] - 1)
        toplam = toplam + gliste1[isindex[sayacindex]]
        nehgantsira.append(toplam)
        sayacindex += 1
        # ilk makinenin gant aşaması tamamlandı
    while sayac11 < len(Neh_gant):
        sayfa5.write(yazmasirasi, sayac11, nehgantsira[sayac11])
        sayac11 += 1
    sayac11 = 0
    yazmasirasi += 1
    # print(glistesonuc)
    while sayacgant < makine_sayisi + 2:
        for cell in sayfa[sayacgant]:
            if cell.value is not None:
                gliste2.append(cell.value)
        gliste2.pop(0)
        while sayac9 < len(Neh_gant):
            while nehgantsira[sayac9] > bekleme:
                bekleme = bekleme + 1
            zaman = bekleme + gliste2[isindex[sayac9]]
            sonliste.append(zaman)
            guncelleme = sonliste[sayac9 - 1] + gliste2[isindex[sayac9]]
            if sonliste[sayac9] < guncelleme:
                sonliste[sayac9] = guncelleme
            if sayac9 == 0:
                sonliste[sayac9] = zaman
            sayac9 += 1
            zaman = 0
            bekleme = 0
        while sayac10 < len(sonliste):
            nehgantsira[sayac10] = sonliste[sayac10]
            sayac10 += 1
        while sayac11 < len(Neh_gant):
            sayfa5.write(yazmasirasi, sayac11, nehgantsira[sayac11])
            sayac11 += 1
        yazmasirasi += 1
        sayac11 = 0
        sonliste.clear()
        sayac9 = 0
        sayac10 = 0
        gliste2.clear()
        sayacgant += 1
    sayacgant = 0
    guncelleme = 0


def SA(yazmasirasi):

    gliste1 = []
    gliste2 = []
    global Sa_fmax
    Sa_fmax = []
    isindex = []
    sonliste = []
    toplam = 0
    sayacindex = 0
    bekleme = 0
    zaman = 0
    guncelleme = 0
    sayac9 = 0
    sayac10 = 0
    sayacgant = 3
    bitis = makine_sayisi + 1

    # işleri index islemine düzeltme
    for cell in sayfa['2']:
        if cell.value is not None:
            gliste1.append(cell.value)
    gliste1.pop(0)
    while sayacindex < len(Simulated_sira):
        isindex.append(Simulated_sira[sayacindex] - 1)
        toplam = toplam + gliste1[isindex[sayacindex]]
        Sa_fmax.append(toplam)
        sayacindex += 1
        # ilk makinenin gant aşaması tamamlandı
    # print(glistesonuc)
    while sayacgant < makine_sayisi + 2:
        for cell in sayfa[sayacgant]:
            if cell.value is not None:
                gliste2.append(cell.value)
        gliste2.pop(0)
        while sayac9 < len(Simulated_sira):
            while Sa_fmax[sayac9] > bekleme:
                bekleme = bekleme + 1
            zaman = bekleme + gliste2[isindex[sayac9]]
            sonliste.append(zaman)
            guncelleme = sonliste[sayac9 - 1] + gliste2[isindex[sayac9]]
            if sonliste[sayac9] < guncelleme:
                sonliste[sayac9] = guncelleme
            if sayac9 == 0:
                sonliste[sayac9] = zaman
            sayac9 += 1
            zaman = 0
            bekleme = 0
        while sayac10 < len(sonliste):
            Sa_fmax[sayac10] = sonliste[sayac10]
            sayac10 += 1
        sayac11 = 0
        sonliste.clear()
        sayac9 = 0
        sayac10 = 0
        gliste2.clear()
        sayacgant += 1
    sayacgant = 0
    guncelleme = 0
    # print(Sa_fmax)


def girdiac():
    os.system("Girdiler.xlsx")


def calistir():
    print("Çalışıyor...")
    window.destroy()


kitap = load_workbook('cds 5x6.xlsx')
kitap1 = xlsxwriter.Workbook('sonuclarrr.xlsx')
window = tk.Tk()
window.geometry('1200x700+10+10')
window.configure(bg='black')
window.resizable(False, False)
# form.state('normal')  #normal ekran
# form.state('zoomed') #tam ekran
window.title('Akış Tipi Atöyler İçin Sıralama Optimizasyonu')
etiket1 = tk.Label(window, text='Makine (M) Sayısı:', font=96, width=30)
etiket2 = tk.Label(window, text='İş (Job) Sayısı:', font=96, width=30)

makinesayisi = Entry(window, justify='right', fg='blue', bg='yellow', font=128, width=5)
issayisi = Entry(window, justify='right', fg='blue', bg='yellow', font=128, width=5)
adim1 = tk.Label(window, text='1. ADIM: Girdiler Excel dosyasını yandaki', font=24)
adim1.place(x=10, y=10)
adim1_2 = tk.Label(window, text='formatta doldurup kaydediniz.', font=24)
adim1_2.place(x=10, y=40)
adim2 = tk.Label(window, text='2. ADIM: n/m/F/Fmax için Makine Sayısı (m) ve ', font=24)
adim2.place(x=10, y=150)
adim2_1 = tk.Label(window, text='İş Sayısı (n) girip "İŞLE ve GETİR" butonuna tıklayınız.', font=24)
adim2_1.place(x=10, y=180)
sonuc1 = tk.Label(window, text='İşlenen Değerler--->', font=24)
sonuc1.place(x=530, y=350)
adim3 = tk.Label(window, text='3. ADIM: Algoritmayı Başlatınız.', font=24)
adim3.place(x=10, y=470)
foto = PhotoImage(file="sablon.png")
sablon = Label(window, image=foto)
sablon.place(x=530, y=10)

etiket1.place(x=10, y=270)
etiket2.place(x=10, y=340)
makinesayisi.place(x=380, y=270)
issayisi.place(x=380, y=340)
buton_isle = Button(window, text='İŞLE VE GETİR', font=128, bg='turquoise', command=degergetir)
buton_isle.place(x=220, y=400)
buton_girdiler = Button(window, text='Dosyayı Aç', font=128, bg='turquoise', command=girdiac)
buton_girdiler.place(x=30, y=80)
check_girdiler = Checkbutton(window, text='Girişleri Yaptım ve Kaydettim.', bg='turquoise', font=72,
                             activebackground='red', state=NORMAL)
check_girdiler.place(x=190, y=85)
buton_baslat = Button(window, text='ÇALIŞTIR', font=128, bg='turquoise', fg='Red', command=calistir)
buton_baslat.place(x=330, y=460)

window.mainloop()

sayfa = kitap.active

sayfa6 = kitap1.add_worksheet('SONUC')
sayfa1 = kitap1.add_worksheet('CDS')
sayfa2 = kitap1.add_worksheet('PALMER')
sayfa3 = kitap1.add_worksheet('RAP')
sayfa4 = kitap1.add_worksheet('GUPTA')
sayfa5 = kitap1.add_worksheet('NEH')


renklendirme = kitap1.add_format()
renklendirme.set_bold(bold=True)
renklendirme.set_bg_color('#FF9999')
renklendirme2 = kitap1.add_format()
renklendirme2.set_bg_color('#CCFFFF')
renklendirme2.set_bold(bold=True)
renklendirme2.set_font_color('#660000')
renklendirme3 = kitap1.add_format()
renklendirme3.set_bg_color('#A0A0A0')
renklendirme3.set_font_size(9)
renklendirme1 = kitap1.add_format()
renklendirme1.set_bg_color('#000066')
renklendirme1.set_font_color('yellow')
renklendirme1.set_bold(bold=True)

sonuc_sira = []
sonuc_fmax = []
yazmasirasi_2 = 1
liste1 = []
liste2 = []
liste3 = []
liste4 = []
liste5 = []
liste6 = []
list1 = []
list2 = []
list3 = []
list4 = []
sonuc = []
sayfano = 0

birlestirveortala = kitap1.add_format({'align': 'center'})
merge_format = kitap1.add_format({'bold': True, 'align': 'center', 'fg_color': '#D7E4BC'})

sayfa6.write(0, 0, 'ALGORİTMA', renklendirme3)
sayfa6.write(0, 1, 'FMAX', renklendirme1)
sayfa6.merge_range('C1:M1', '>>>>>>>>  İ Ş    S I R A L A M A S I >>>>>>>>', merge_format)
is_sayisi = ssayisi
makine_sayisi = msayisi
# CDS CDS CDS CDS CDS CDS CDS CDS CDS CDS CDS CDS CDS CDS CDS CDS CDS CDS CDS CDS CDS CDS CDS CDS CDS CDS CDS CDS CDS
for cell in sayfa['2']:
    if cell.value is not None:
        liste1.append(cell.value)
liste1.pop(0)

for cell in sayfa[makine_sayisi + 1]:
    if cell.value is not None:
        liste2.append(cell.value)
liste2.pop(0)

sayac = 0
islistesi = []
islistesi1 = []
sayac1 = 0
sayac3 = 0
sayac5 = 0
sayac6 = 0
sayac7 = 0
sayac11 = 0
yazmasirasi = 0
islemsayisi = makine_sayisi - 2
islemsayac = 0
secimislistesi = []
secimislistesi1 = []
secilmeyenislistesi = []
secilmeyenislistesi1 = []
secilmeyenlistesi1 = []
secilmeyenlistesi2 = []
secilmeyenlistesi3 = []
secilmeyenlistesi4 = []
baslangic = 2

# sıfırdan bire iş listesi eklemesi yapılır#
while sayac1 < is_sayisi:
    sayac1 += 1
    islistesi.append(sayac1)

while is_sayisi > sayac:
    if min(liste1) < liste2[liste1.index(min(liste1))] or min(liste1) == liste2[liste1.index(min(liste1))]:
        secimislistesi.append(islistesi[liste1.index(min(liste1))])
        del islistesi[liste1.index(min(liste1))]
        del liste2[liste1.index(min(liste1))]
        del liste1[liste1.index(min(liste1))]

    else:
        secilmeyenlistesi1.append(min(liste1))
        secilmeyenlistesi2.append(liste2[liste1.index(min(liste1))])
        secilmeyenislistesi.append(islistesi[liste1.index(min(liste1))])
        del islistesi[liste1.index(min(liste1))]
        del liste2[liste1.index(min(liste1))]
        del liste1[liste1.index(min(liste1))]

    sayac += 1

while len(secilmeyenlistesi2) > 0:
    secimislistesi.append(secilmeyenislistesi[secilmeyenlistesi2.index(max(secilmeyenlistesi2))])
    secilmeyenislistesi.remove(secilmeyenislistesi[secilmeyenlistesi2.index(max(secilmeyenlistesi2))])
    secilmeyenlistesi2.remove(max(secilmeyenlistesi2))

secilmeyenlistesi1.clear()
sonuc_sira.append(secimislistesi)
# print(secimislistesi)
while sayac11 < is_sayisi:
    sayfa1.write(yazmasirasi, sayac11, secimislistesi[sayac11], renklendirme2)
    sayac11 += 1
sayac11 = 0
yazmasirasi += 1
gliste1 = []
gliste2 = []
glistesonuc = []
isindex = []
sonliste = []
toplam = 0
sayacindex = 0
bekleme = 0
zaman = 0
guncelleme = 0
sayac9 = 0
sayac10 = 0
sayacgant = 3
bitis = makine_sayisi + 1

# işleri index islemine düzeltme
for cell in sayfa['2']:
    if cell.value is not None:
        gliste1.append(cell.value)
gliste1.pop(0)
while sayacindex < is_sayisi:
    isindex.append(secimislistesi[sayacindex] - 1)
    toplam = toplam + gliste1[isindex[sayacindex]]
    glistesonuc.append(toplam)
    sayacindex += 1
    # ilk makinenin gant aşaması tamamlandı
while sayac11 < is_sayisi:
    sayfa1.write(yazmasirasi, sayac11, glistesonuc[sayac11])
    sayac11 += 1
sayac11 = 0
yazmasirasi += 1
# print(glistesonuc)
while sayacgant < makine_sayisi + 2:
    for cell in sayfa[sayacgant]:
        if cell.value is not None:
            gliste2.append(cell.value)
    gliste2.pop(0)
    while sayac9 < is_sayisi:
        while glistesonuc[sayac9] > bekleme:
            bekleme = bekleme + 1
        zaman = bekleme + gliste2[isindex[sayac9]]
        sonliste.append(zaman)
        guncelleme = sonliste[sayac9 - 1] + gliste2[isindex[sayac9]]
        if sonliste[sayac9] < guncelleme:
            sonliste[sayac9] = guncelleme
        if sayac9 == 0:
            sonliste[sayac9] = zaman
        sayac9 += 1
        zaman = 0
        bekleme = 0
    while sayac10 < len(sonliste):
        glistesonuc[sayac10] = sonliste[sayac10]
        sayac10 += 1
    while sayac11 < is_sayisi:
        sayfa1.write(yazmasirasi, sayac11, glistesonuc[sayac11])
        sayac11 += 1
    yazmasirasi += 1
    sayac11 = 0
    sonliste.clear()
    # print(glistesonuc)
    sayac9 = 0
    sayac10 = 0
    gliste2.clear()
    sayacgant += 1
sayacgant = 0
guncelleme = 0
sonuc_fmax.append(max(glistesonuc))
sayfa6.write(yazmasirasi_2, 0, 'CDS', renklendirme)
sayfa6.write(yazmasirasi_2, 1, sonuc_fmax[0], renklendirme1)
sayfa6.write_row(yazmasirasi_2, 2, sonuc_sira[0])
yazmasirasi_2 += 1
sonuc_sira.clear()
sonuc_fmax.clear()
# ilk siralama islemleri bitti
for cell in sayfa[baslangic]:
    if cell.value is not None:
        list1.append(cell.value)
list1.pop(0)

for cell in sayfa[makine_sayisi + 1]:
    if cell.value is not None:
        list2.append(cell.value)
list2.pop(0)

while islemsayac < islemsayisi:
    for cell in sayfa[baslangic + 1]:
        if cell.value is not None:
            list3.append(cell.value)
    list3.pop(0)

    for cell in sayfa[bitis - 1]:
        if cell.value is not None:
            list4.append(cell.value)
    list4.pop(0)
    while sayac3 < is_sayisi:
        list1[sayac3] = list1[sayac3] + list3[sayac3]
        list2[sayac3] = list2[sayac3] + list4[sayac3]
        sayac3 += 1
    while sayac7 < is_sayisi:
        liste5.append(list1[sayac7])
        liste6.append(list2[sayac7])
        sayac7 += 1
    sayac7 = 0
    baslangic += 1
    bitis -= 1
    list4.clear()
    list3.clear()
    sayac3 = 0
    islemsayac += 1
    while sayac5 < is_sayisi:
        sayac5 += 1
        islistesi1.append(sayac5)
    sayac5 = 0
    while is_sayisi > sayac6:
        if min(liste5) < liste6[liste5.index(min(liste5))] or min(liste5) == liste6[liste5.index(min(liste5))]:
            secimislistesi1.append(islistesi1[liste5.index(min(liste5))])
            del islistesi1[liste5.index(min(liste5))]
            del liste6[liste5.index(min(liste5))]
            del liste5[liste5.index(min(liste5))]

        else:
            secilmeyenlistesi4.append(liste6[liste5.index(min(liste5))])
            secilmeyenlistesi3.append(liste5[liste5.index(min(liste5))])
            secilmeyenislistesi1.append(islistesi1[liste5.index(min(liste5))])
            del liste6[liste5.index(min(liste5))]
            del islistesi1[liste5.index(min(liste5))]
            del liste5[liste5.index(min(liste5))]

        sayac6 += 1

    sayac6 = 0
    while len(secilmeyenlistesi4) > 0:
        secimislistesi1.append(secilmeyenislistesi1[secilmeyenlistesi4.index(max(secilmeyenlistesi4))])
        secilmeyenislistesi1.remove(secilmeyenislistesi1[secilmeyenlistesi4.index(max(secilmeyenlistesi4))])
        secilmeyenlistesi4.remove(max(secilmeyenlistesi4))
    secilmeyenlistesi3.clear()
    while sayac11 < is_sayisi:
        sayfa1.write(yazmasirasi, sayac11, secimislistesi1[sayac11], renklendirme2)
        sonuc_sira.append(secimislistesi1[sayac11])
        sayac11 += 1
    yazmasirasi += 1
    gliste1 = []
    gliste2 = []
    glistesonuc = []
    isindex = []
    sonliste = []
    toplam = 0
    sayacindex = 0
    bekleme = 0
    zaman = 0
    sayac9 = 0
    sayac10 = 0
    sayacgant = 3
    sayac11 = 0
    for cell in sayfa['2']:
        if cell.value is not None:
            gliste1.append(cell.value)
    gliste1.pop(0)
    while sayacindex < is_sayisi:
        isindex.append(secimislistesi1[sayacindex] - 1)
        toplam = toplam + gliste1[isindex[sayacindex]]
        glistesonuc.append(toplam)
        sayacindex += 1
        # ilk makinenin gant asamasi tamamlandi
    while sayac11 < is_sayisi:
        sayfa1.write(yazmasirasi, sayac11, glistesonuc[sayac11])
        sayac11 += 1
    sayac11 = 0
    yazmasirasi += 1
    # print(glistesonuc)
    while sayacgant < makine_sayisi + 2:
        for cell in sayfa[sayacgant]:
            if cell.value is not None:
                gliste2.append(cell.value)
        gliste2.pop(0)
        while sayac9 < is_sayisi:
            while glistesonuc[sayac9] > bekleme:
                bekleme = bekleme + 1
            zaman = bekleme + gliste2[isindex[sayac9]]
            sonliste.append(zaman)
            guncelleme = sonliste[sayac9 - 1] + gliste2[isindex[sayac9]]
            if sonliste[sayac9] < guncelleme:
                sonliste[sayac9] = guncelleme
            if sayac9 == 0:
                sonliste[sayac9] = zaman
            sayac9 += 1
            zaman = 0
            bekleme = 0
        while sayac10 < len(sonliste):
            glistesonuc[sayac10] = sonliste[sayac10]
            sayac10 += 1
        while sayac11 < is_sayisi:
            sayfa1.write(yazmasirasi, sayac11, glistesonuc[sayac11])
            sayac11 += 1
        yazmasirasi += 1
        sayac11 = 0
        sonliste.clear()
        # print(glistesonuc)
        sayac9 = 0
        sayac10 = 0
        gliste2.clear()
        sayacgant += 1
    sayacgant = 0
    guncelleme = 0
    sonuc_fmax.append(max(glistesonuc))
    secimislistesi1.clear()
    secilmeyenlistesi4.clear()
    secilmeyenlistesi3.clear()
    secilmeyenislistesi1.clear()
    islistesi1.clear()
    sayfa6.write(yazmasirasi_2, 0, 'CDS', renklendirme)
    sayfa6.write(yazmasirasi_2, 1, sonuc_fmax[0], renklendirme1)
    sayfa6.write_row(yazmasirasi_2, 2, sonuc_sira)
    yazmasirasi_2 += 1
    sonuc_sira.clear()
    sonuc_fmax.clear()

# PALMER  PALMER PALMER PALMER PALMER PALMER PALMER PALMER PALMER PALMER PALMER PALMER PALMER PALMER PALMER
yazmasirasi = 0
sayac2 = 0
sayac4 = 0
sayac8 = 0
sayac12 = 2
sayac13 = 0
baslangic = 2
mliste = []
psonuc = []
plistesi = []
lgirdiler = []
palmer = []
psiralama = []
while sayac4 < is_sayisi:
    sayac4 += 1
    plistesi.append(sayac4)
while sayac8 < makine_sayisi:
    sayac8 += 1
    mliste.append(sayac8)
while sayac12 < is_sayisi + 2:
    for column in range(sayac12, sayac12 + 1):
        for row in range(2, makine_sayisi + 2):
            value = sayfa.cell(row, column).value
            lgirdiler.append(value)
    # print(lgirdiler)
    while sayac2 < makine_sayisi:
        palmer.append(- ((makine_sayisi - ((2 * mliste[sayac2]) - 1)) * lgirdiler[sayac2]))
        # print(palmer)
        sayac2 += 1
    lgirdiler.clear()
    sayac12 += 1
    sayac2 = 0
    psonuc.append(sum(palmer))
    palmer.clear()
    # print(psonuc)
while is_sayisi > sayac13:
    cikartilan = psonuc.index(max(psonuc))
    psiralama.append(plistesi[cikartilan])
    plistesi.remove(plistesi[psonuc.index(max(psonuc))])
    psonuc.remove(psonuc[psonuc.index(max(psonuc))])

    # print("--", psiralama)
    sayac13 += 1
# print(plistesi)
# print(psiralama)

GANT('PALMER', psiralama, sayfa2, yazmasirasi_2)

# RAP RAP RAP RAP RAP RAP RAP RAP RAP RAP RAP RAP RAP RAP RAP RAP RAP RAP RAP RAP RAP RAP RAP RAP RAP RAP RAP RAP RAP
yazmasirasi = 0
sayacindex = 0
sayac13 = 0
sayac14 = 0
sayac15 = 2
sayac16 = 0
sayac17 = 0
rlistesi = []
rgirdiler = []
j_liste = []
jliste = []
rap1 = []
rap2 = []
rapsonuc1 = []
rapsonuc2 = []
rapsonuc = []
while sayac13 < is_sayisi:
    sayac13 += 1
    rlistesi.append(sayac13)

while sayac14 < makine_sayisi:
    sayac14 += 1
    j_liste.append(sayac14)
jliste = j_liste[::-1]
while sayac15 < is_sayisi + 2:
    for columnn in range(sayac15, sayac15 + 1):
        for roww in range(2, makine_sayisi + 2):
            value = sayfa.cell(roww, columnn).value
            rgirdiler.append(value)
    # print(lgirdiler)
    while sayac16 < makine_sayisi:
        rap1.append(jliste[sayac16] * rgirdiler[sayac16])
        rap2.append(j_liste[sayac16] * rgirdiler[sayac16])
        sayac16 += 1

    rgirdiler.clear()
    sayac15 += 1
    sayac16 = 0
    rapsonuc1.append(sum(rap1))
    rap1.clear()
    rapsonuc2.append(sum(rap2))
    rap2.clear()
# print(rapsonuc1)
# print(rapsonuc2)

sayac11 = 0
gliste1 = []
gliste2 = []
glistesonuc = []
isindex = []
sonliste = []
toplam = 0
sayacindex = 0
bekleme = 0
zaman = 0
guncelleme = 0
sayac9 = 0
sayac10 = 0
sayacgant = 3

while is_sayisi > sayac17:
    if min(rapsonuc1) < rapsonuc2[rapsonuc1.index(min(rapsonuc1))] \
            or min(rapsonuc1) == rapsonuc2[rapsonuc1.index(min(rapsonuc1))]:
        rapsonuc.append(rlistesi[rapsonuc1.index(min(rapsonuc1))])
        del rlistesi[rapsonuc1.index(min(rapsonuc1))]
        del rapsonuc2[rapsonuc1.index(min(rapsonuc1))]
        del rapsonuc1[rapsonuc1.index(min(rapsonuc1))]

    else:
        secilmeyenlistesi1.append(min(rapsonuc1))
        secilmeyenlistesi2.append(rapsonuc2[rapsonuc1.index(min(rapsonuc1))])
        secilmeyenislistesi.append(rlistesi[rapsonuc1.index(min(rapsonuc1))])
        del rlistesi[rapsonuc1.index(min(rapsonuc1))]
        del rapsonuc2[rapsonuc1.index(min(rapsonuc1))]
        del rapsonuc1[rapsonuc1.index(min(rapsonuc1))]

    sayac17 += 1

while len(secilmeyenlistesi2) > 0:
    rapsonuc.append(secilmeyenislistesi[secilmeyenlistesi2.index(max(secilmeyenlistesi2))])
    secilmeyenislistesi.remove(secilmeyenislistesi[secilmeyenlistesi2.index(max(secilmeyenlistesi2))])
    secilmeyenlistesi2.remove(max(secilmeyenlistesi2))

# print(rapsonuc)
yazmasirasi_2 += 1
GANT('RAP', rapsonuc, sayfa3, yazmasirasi_2)

# GUPTA GUPTA GUPTA GUPTA GUPTA GUPTA GUPTA GUPTA GUPTA GUPTA GUPTA GUPTA GUPTA GUPTA GUPTA GUPTA GUPTA GUPTA GUPTA
yazmasirasi = 0
sayac18 = 0
sayac19 = 2
sayac20 = 0
guptakiyas = 0
guptasayi = 0
mingupta = []
g_is_liste = []
g_girdiler = []
gup_sonuc = []
gup_siralama = []
while sayac18 < is_sayisi:
    sayac18 += 1
    g_is_liste.append(sayac18)

while sayac19 < is_sayisi + 2:
    for columnnn in range(sayac19, sayac19 + 1):
        for rowww in range(2, makine_sayisi + 2):
            value = sayfa.cell(rowww, columnnn).value
            g_girdiler.append(value)

    if g_girdiler[makine_sayisi - 1] > g_girdiler[0]:
        guptakiyas = - 1
    else:
        guptakiyas = 1
    while sayac20 < makine_sayisi - 2:
        mingupta.append(g_girdiler[sayac20] + g_girdiler[sayac20 + 1])
        sayac20 += 1
        guptasayi = guptakiyas / min(mingupta)
    # print(guptasayi)
    gup_siralama.append(guptasayi)
    mingupta.clear()
    g_girdiler.clear()
    sayac19 += 1
    sayac20 = 0
while sayac20 < is_sayisi:
    gup_sonuc.append(g_is_liste[gup_siralama.index(min(gup_siralama))])
    del g_is_liste[gup_siralama.index(min(gup_siralama))]
    del gup_siralama[gup_siralama.index(min(gup_siralama))]
    sayac20 += 1
# print(gup_sonuc)
# print(gup_siralama)
yazmasirasi_2 += 1
GANT('GUPTA', gup_sonuc, sayfa4, yazmasirasi_2)

# NEH NEH NEH NEH NEH NEH NEH NEH NEH NEH NEH NEH NEH NEH NEH NEH NEH NEH NEH NEH NEH NEH NEH NEH NEH NEH NEH NEH NEH
sayac20 = 0
sayac19 = 2
sayac18 = 0
Neh_is = []
Neh_girdiler = []
Neh_top_girdi = []
Neh_sira = []
Neh_gant = []
while sayac18 < is_sayisi:
    sayac18 += 1
    Neh_is.append(sayac18)

while sayac19 < is_sayisi + 2:
    for columnnn in range(sayac19, sayac19 + 1):
        for rowww in range(2, makine_sayisi + 2):
            value = sayfa.cell(rowww, columnnn).value
            Neh_girdiler.append(value)
    Neh_top_girdi.append(sum(Neh_girdiler))
    Neh_girdiler.clear()
    # print(Neh_top_girdi)
    sayac19 += 1

while sayac20 < is_sayisi:
    Neh_sira.append(Neh_is[Neh_top_girdi.index(max(Neh_top_girdi))])
    del Neh_is[Neh_top_girdi.index(max(Neh_top_girdi))]
    del Neh_top_girdi[Neh_top_girdi.index(max(Neh_top_girdi))]
    sayac20 += 1
# print(Neh_sira)

Neh_gant.append(Neh_sira[0])
del Neh_sira[0]
Neh_gant.append(Neh_sira[0])
del Neh_sira[0]

yazmasirasi = 0

NEH(yazmasirasi)
ilkNeh = nehgantsira[1]
Neh_gant.reverse()
NEH(yazmasirasi)
if ilkNeh < nehgantsira[1]:
    Neh_gant.reverse()
    NEH(yazmasirasi)
yazmasirasi_4 = makine_sayisi + 3
sayac21 = 0
sayac22 = 0
exitdongu = -2
kontrolsonuc = []
sayfa5.write(makine_sayisi + 2, 0, 'NEH Sıralama Seçim Sonuçları')
while sayac21 < is_sayisi - 2:

    while sayac22 < sayac21 + 3 and exitdongu == -2:
        Neh_gant.insert(sayac22, Neh_sira[0])
        NEH(yazmasirasi)
        kontrolsonuc.append(nehgantsira[len(Neh_gant) - 1])

        sayac24 = 0

        while sayac24 < len(Neh_gant):
            sayfa5.write(yazmasirasi_4, 0, kontrolsonuc[sayac22], renklendirme)
            sayfa5.write(yazmasirasi_4, sayac24 + 1, Neh_gant[sayac24])
            sayac24 += 1
        yazmasirasi_4 += 1

        del Neh_gant[sayac22]

        if sayac22 == sayac21 + 2:
            # print(sayac22, Neh_gant, kontrolsonuc)
            sayac22 = kontrolsonuc.index(min(kontrolsonuc))
            Neh_gant.insert(sayac22, Neh_sira[0])
            NEH(yazmasirasi)
            exitdongu = -1
        sayac22 += 1
    exitdongu = -2
    # print("bitis", Neh_gant)
    kontrolsonuc.clear()
    del Neh_sira[0]
    sayac22 = 0
    sayac21 += 1
    yazmasirasi_4 += 1
# print(Neh_gant)
yazmasirasi_2 += 1
sayfa6.write(yazmasirasi_2, 0, 'NEH', renklendirme)
sayfa6.write(yazmasirasi_2, 1, nehgantsira[is_sayisi - 1], renklendirme1)
sayac23 = 0
while sayac23 < is_sayisi:
    sayfa6.write(yazmasirasi_2, sayac23 + 2, Neh_gant[sayac23])
    sayac23 += 1

kitap1.close()


# os.system("sonuclarrr.xlsx")

window_3 = tk.Tk()
window_3.geometry('1300x400+150+150')
window_3.configure(bg='black')
window_3.resizable(True, False)
# form.state('normal')  #normal ekran
# form.state('zoomed') #tam ekran
window_3.title(' Tavlama Benzetimi (Simulated Annealing)')
buton_sonuc = Button(window_3, text='Sonuçları Göster', font=128, bg='turquoise', command=sonucac)
buton_sonuc.place(x=150, y=120)
# buton_tavlama = Button(window_3, text='TAVLAMA BENZETİMİ BAŞLAT', font=128, bg='red')


def tumexcel():
    os.system('yapaysonuc.xlsx')


buton_tavlama = Button(window_3, text='Tüm İhtimalleri gör', font=128, bg='red', command=tumexcel)
buton_tavlama.place(x=450, y=120)

etiket3 = Label(window_3, text='EN İYİ FMAX', fg='orange', bg='GRAY', width=12, font=24)
etiket3.place(x=10, y=10)
etiket4 = Label(window_3, text='EN İYİ SIRALAMA', fg='orange', bg='GRAY', width=100, font=18, wraplength=1000)
etiket4.place(x=180, y=10)

kitap2 = load_workbook('sonuclarrr.xlsx')
sayfa_sonuc = kitap2.active


sonuc_siralamasi = []
fmax_liste = []

for columnnnn in range(2, 3):
    for rowwww in range(2, makine_sayisi + 5):
        value = sayfa_sonuc.cell(rowwww, columnnnn).value
        if value is not None:
            fmax_liste.append(value)

for cell in sayfa_sonuc[fmax_liste.index(min(fmax_liste)) + 2]:
    if cell.value is not None:
        sonuc_siralamasi.append(cell.value)
del sonuc_siralamasi[0]
iyiFmax = sonuc_siralamasi[0]
del sonuc_siralamasi[0]
# del sonuc_siralamasi[0]

# print(fmax_liste)
# print(iyiFmax)
kitap2.close()
kitap3 = xlsxwriter.Workbook('yapaysonuc.xlsx')

sayfa7 = kitap3.add_worksheet('SA')

bas_sonuc1 = Label(window_3, text=iyiFmax, fg='black', bg='yellow', width=12, font=32)
bas_sonuc1.place(x=10, y=40)
bas_sonuc = Label(window_3, text=sonuc_siralamasi, fg='orange', bg='white', width=100, font=24, wraplength=1000)
bas_sonuc.place(x=180, y=40)

# Tavlama Benzetimi # Tavlama Benzetimi # Tavlama Benzetimi # Tavlama Benzetimi # Tavlama Benzetimi # Tavlama Benzetimi
iterasyon = 0
Simulated_sira = []
# sayac25 = 0
# yazmasirasi_3 = 0
# yazmasirasi_5 = 0
#
# # Tüm olasılık hesabı
# perm = permutations(sonuc_siralamasi)
# sayac26 = 0
#
# for i in list(perm):
#     # print(i)
#     while sayac26 < is_sayisi:
#         Simulated_sira.append(i[sayac26])
#         sayfa7.write(yazmasirasi_5, sayac26, i[sayac26], renklendirme1)
#         sayac26 += 1
#     sayac26 = 0
#     yazmasirasi_5 += 1
#
#     SA(yazmasirasi)
#     print(Simulated_sira)
#     sayfa7.write(yazmasirasi_3, len(sonuc_siralamasi), Sa_fmax[len(Simulated_sira) - 1], renklendirme1)
#     yazmasirasi_3 += 1
#     Simulated_sira.clear()
# # Tüm olasılık hesabı

# yenilemeli iterasyon
# while iterasyon < 10:
#    iterasyon += 1
#    random.shuffle(Simulated_sira)
#    SA(yazmasirasi)
#
#
# # Başlangıç Çözümü
# olasiliklar = []
# while sayac25 < len(sonuc_siralamasi):
#     Simulated_sira.append(sonuc_siralamasi[sayac25])
#     sayac25 += 1
# perm = permutations(Simulated_sira)
# sayac26 = 0
# for i in list(perm):
#     # print(i)
#     while sayac26 < is_sayisi:
#         olasiliklar.append(i[sayac26])
#         sayac26 += 1
#     sayac26 = 0
# print(olasiliklar)
# while iterasyon < 10:
#    iterasyon += 1
#    random.shuffle(Simulated_sira)
#    SA(yazmasirasi)
# yenilemeli iterasyon sonu

# sonuc_siralamasi   ---> algoritmalar sonucu en iyi sıralama (başlangıç sıralaması)
# SA(yazmasirasi)
# Sa_fmax ----> FMax değeri yani SA(yazmasirasi) fonksiyonun bulduğu değer. En küçükleme yapmak istediğimiz. 53'e yaklaşmalı
# Simulated_sira ------>  fonksiyonda hesaplanacak sıra. aynı zamanda move ile yeri değiştirilecek sira

print()

kitap3.close()
window_3.mainloop()
