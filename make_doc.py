from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

OUT = r'C:\Users\Administrator\Documents\N\Laporan_Proyek_Sistem_Perhitungan_Nilai_Akademik.docx'
doc = Document()
section = doc.sections[0]
for side in ('top_margin', 'bottom_margin', 'left_margin', 'right_margin'):
    setattr(section, side, Cm(3))
style = doc.styles['Normal']
style.font.name = 'Times New Roman'
style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
style.font.size = Pt(12)
style.paragraph_format.line_spacing = 1.5

def set_run(run, bold=False):
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    run.font.size = Pt(12)
    run.bold = bold

def paragraph(text='', align=None, indent=True, bold=False):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(0)
    if indent:
        p.paragraph_format.first_line_indent = Cm(1.25)
    if align is not None:
        p.alignment = align
    set_run(p.add_run(text), bold)
    return p

def heading(text, chapter=False):
    p = paragraph(text, WD_ALIGN_PARAGRAPH.CENTER if chapter else None, False, True)
    p.paragraph_format.space_before = Pt(12 if chapter else 6)
    p.paragraph_format.keep_with_next = True

def numbered(items):
    for number, text in enumerate(items, 1):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.left_indent = Cm(1.25)
        p.paragraph_format.first_line_indent = Cm(-0.6)
        set_run(p.add_run(f'{number}. {text}'))

def make_table(headers, rows):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, text in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ''
        set_run(cell.paragraphs[0].add_run(text), True)
        shade = OxmlElement('w:shd')
        shade.set(qn('w:fill'), 'D9EAF7')
        cell._tc.get_or_add_tcPr().append(shade)
    for row in rows:
        cells = table.add_row().cells
        for i, text in enumerate(row):
            cells[i].text = ''
            p = cells[i].paragraphs[0]
            p.paragraph_format.line_spacing = 1.5
            set_run(p.add_run(str(text)))
    doc.add_paragraph()

# Halaman sampul
for _ in range(5): doc.add_paragraph()
paragraph('LAPORAN PROYEK', WD_ALIGN_PARAGRAPH.CENTER, False, True)
paragraph('APLIKASI WEB SISTEM PERHITUNGAN NILAI AKADEMIK SISWA', WD_ALIGN_PARAGRAPH.CENTER, False, True)
for _ in range(7): doc.add_paragraph()
for text in [
    'Disusun oleh : ................................................',
    'NIM          : ................................................',
    'Program Studi: Pendidikan Teknologi Informasi',
    'Semester     : 1 (Satu)',
    'Mata Kuliah  : ................................................',
    'Universitas  : ................................................'
]:
    paragraph(text, WD_ALIGN_PARAGRAPH.CENTER, False)
for _ in range(5): doc.add_paragraph()
paragraph('2026', WD_ALIGN_PARAGRAPH.CENTER, False, True)
doc.add_page_break()

heading('BAB I\nPENDAHULUAN', True)
heading('1.1 Latar Belakang')
paragraph('Pada pembelajaran di sekolah, guru perlu mengolah nilai siswa dari beberapa komponen, seperti tugas, UTS, UAS, dan kehadiran. Nilai tersebut kemudian dihitung menjadi nilai akhir untuk menentukan predikat dan status kelulusan. Jika proses ini dilakukan secara manual, guru atau pengguna dapat membutuhkan waktu lebih lama dan berisiko salah menghitung.')
paragraph('Sebagai mahasiswa semester 1 Program Studi Pendidikan Teknologi Informasi, proyek ini dibuat untuk menerapkan dasar pemrograman web yang telah dipelajari. Aplikasi dibuat menggunakan HTML, CSS, dan JavaScript. Melalui proyek ini, penulis belajar membuat form, mengolah input pengguna, menggunakan percabangan, menyimpan data, serta menampilkan hasil perhitungan pada halaman web.')
heading('1.2 Rumusan Masalah')
numbered(['Bagaimana membuat form untuk memasukkan data nilai siswa?', 'Bagaimana menghitung nilai akhir dari nilai tugas, UTS, UAS, dan kehadiran?', 'Bagaimana menentukan predikat dan status kelulusan secara otomatis?', 'Bagaimana menyimpan data nilai agar tidak langsung hilang saat halaman dimuat ulang?'])
heading('1.3 Tujuan Proyek')
numbered(['Membuat aplikasi web sederhana untuk menghitung nilai akademik siswa.', 'Menerapkan materi dasar HTML, CSS, dan JavaScript.', 'Menampilkan nilai akhir, predikat, dan status kelulusan secara otomatis.', 'Melatih kemampuan mahasiswa dalam membuat proyek web sederhana.'])
heading('1.4 Manfaat Proyek')
numbered(['Memudahkan proses perhitungan nilai siswa.', 'Mengurangi kesalahan hitung yang dapat terjadi saat menghitung secara manual.', 'Menjadi latihan penerapan materi pemrograman web dasar.', 'Menambah pemahaman mahasiswa tentang penggunaan localStorage pada browser.'])
doc.add_page_break()

heading('BAB II\nDESKRIPSI DAN PERANCANGAN APLIKASI', True)
heading('2.1 Deskripsi Aplikasi')
paragraph('Aplikasi Sistem Perhitungan Nilai Akademik Siswa adalah aplikasi web satu halaman. Pengguna mengisi nama siswa, NIS, nilai tugas, UTS, UAS, dan kehadiran. Setelah tombol simpan ditekan, sistem memeriksa data, menghitung nilai akhir, menentukan predikat, lalu menentukan apakah siswa lulus atau tidak lulus.')
paragraph('Data yang sudah dimasukkan ditampilkan pada tabel hasil. Aplikasi juga menyediakan informasi jumlah siswa, rata-rata nilai kelas, dan jumlah siswa yang lulus. Data tersimpan pada localStorage browser sehingga masih dapat ditampilkan saat halaman dibuka kembali di browser yang sama.')
heading('2.2 Kebutuhan Perangkat')
make_table(['Jenis', 'Kebutuhan'], [
    ['Perangkat keras', 'Laptop atau komputer untuk menjalankan browser.'],
    ['Perangkat lunak', 'Browser seperti Google Chrome, Microsoft Edge, atau Mozilla Firefox.'],
    ['Bahasa pemrograman', 'HTML, CSS, dan JavaScript.'],
    ['Penyimpanan data', 'localStorage pada browser.']
])
heading('2.3 Fitur Aplikasi')
make_table(['No.', 'Fitur', 'Penjelasan'], [
    ['1', 'Input data siswa', 'Mengisi nama, NIS, dan empat nilai siswa.'],
    ['2', 'Validasi data', 'Memastikan semua data terisi dan nilai berada pada rentang 0 sampai 100.'],
    ['3', 'Cek NIS', 'Menolak NIS yang sudah pernah dimasukkan.'],
    ['4', 'Hitung nilai akhir', 'Menghitung nilai berdasarkan bobot tiap komponen.'],
    ['5', 'Predikat dan status', 'Menentukan huruf predikat serta status Lulus atau Tidak Lulus.'],
    ['6', 'Statistik kelas', 'Menampilkan jumlah siswa, rata-rata kelas, dan jumlah lulus.'],
    ['7', 'Hapus data', 'Menghapus data siswa setelah pengguna menyetujui konfirmasi.']
])
heading('2.4 Alur Kerja Aplikasi')
numbered(['Pengguna mengisi data siswa pada form.', 'Sistem memeriksa kelengkapan data, rentang nilai, dan NIS duplikat.', 'Sistem menghitung nilai akhir.', 'Sistem menentukan predikat dan status kelulusan.', 'Data disimpan ke localStorage.', 'Tabel dan statistik kelas diperbarui.'])
doc.add_page_break()

heading('BAB III\nIMPLEMENTASI PROGRAM', True)
heading('3.1 Teknologi yang Digunakan')
make_table(['Teknologi', 'Fungsi dalam Proyek'], [
    ['HTML', 'Membuat struktur halaman, form input, tabel, tombol, dan bagian statistik.'],
    ['CSS', 'Mengatur warna, ukuran, layout, kartu, tabel, tombol, serta tampilan responsif.'],
    ['JavaScript', 'Mengolah input, validasi, perhitungan, penentuan predikat, dan pembaruan tampilan.'],
    ['localStorage', 'Menyimpan data nilai pada browser pengguna.']
])
heading('3.2 Rumus Nilai Akhir')
paragraph('Aplikasi menggunakan empat komponen penilaian. Bobot nilai dapat dilihat pada tabel berikut.', False)
make_table(['Komponen', 'Bobot'], [['Tugas', '30%'], ['UTS', '30%'], ['UAS', '30%'], ['Kehadiran', '10%']])
paragraph('Nilai Akhir = (Tugas × 0,30) + (UTS × 0,30) + (UAS × 0,30) + (Kehadiran × 0,10)', WD_ALIGN_PARAGRAPH.CENTER, False, True)
heading('3.3 Aturan Predikat dan Kelulusan')
make_table(['Rentang Nilai', 'Predikat', 'Keterangan'], [['85–100', 'A', 'Sangat Baik'], ['75–84,99', 'B', 'Baik'], ['65–74,99', 'C', 'Cukup'], ['50–64,99', 'D', 'Kurang'], ['0–49,99', 'E', 'Sangat Kurang']])
paragraph('KKM pada aplikasi ini adalah 75. Siswa dengan nilai akhir minimal 75 mendapat status Lulus. Siswa dengan nilai akhir di bawah 75 mendapat status Tidak Lulus.', False)
heading('3.4 Penjelasan Bagian JavaScript')
make_table(['Bagian', 'Penjelasan Sederhana'], [
    ['STORAGE_KEY', 'Nama kunci untuk menyimpan data pada localStorage.'],
    ['muatData()', 'Mengambil data nilai yang sebelumnya tersimpan di browser.'],
    ['simpanData()', 'Menyimpan data siswa ke localStorage dalam bentuk JSON.'],
    ['form.addEventListener()', 'Menjalankan proses saat tombol simpan pada form ditekan.'],
    ['tampilkanData()', 'Menampilkan semua data siswa pada tabel.'],
    ['escapeHTML()', 'Mengubah karakter tertentu agar teks input aman saat ditampilkan.'],
    ['hapusData()', 'Menghapus data siswa yang dipilih pengguna.'],
    ['hitungStatistik()', 'Menghitung jumlah siswa, rata-rata kelas, dan jumlah siswa lulus.']
])
doc.add_page_break()

heading('BAB IV\nPENGUJIAN SEDERHANA', True)
heading('4.1 Pengujian Fitur')
make_table(['No.', 'Yang Diuji', 'Hasil yang Diharapkan'], [
    ['1', 'Semua input diisi benar', 'Data dapat disimpan dan hasil perhitungan muncul pada tabel.'],
    ['2', 'Nilai kurang dari 0 atau lebih dari 100', 'Sistem menampilkan peringatan dan data tidak disimpan.'],
    ['3', 'Nama atau NIS kosong', 'Sistem meminta pengguna melengkapi data.'],
    ['4', 'NIS sama', 'Sistem menampilkan pesan bahwa NIS sudah terdaftar.'],
    ['5', 'Nilai akhir 75 atau lebih', 'Status yang ditampilkan adalah Lulus.'],
    ['6', 'Nilai akhir kurang dari 75', 'Status yang ditampilkan adalah Tidak Lulus.'],
    ['7', 'Halaman dimuat ulang', 'Data tetap muncul selama localStorage browser belum dihapus.'],
    ['8', 'Tombol Hapus dipilih', 'Data dihapus setelah pengguna menekan setuju pada konfirmasi.']
])
heading('4.2 Contoh Perhitungan')
paragraph('Contoh nilai siswa adalah tugas 80, UTS 75, UAS 90, dan kehadiran 100. Perhitungannya sebagai berikut.', False)
paragraph('Nilai Akhir = (80 × 0,30) + (75 × 0,30) + (90 × 0,30) + (100 × 0,10) = 83,50', WD_ALIGN_PARAGRAPH.CENTER, False, True)
paragraph('Nilai akhir 83,50 berada pada rentang 75–84,99. Jadi, siswa mendapatkan predikat B atau Baik. Karena nilainya lebih dari KKM 75, status siswa adalah Lulus.', False)
heading('4.3 Hasil Pengujian')
paragraph('Berdasarkan pengujian sederhana, aplikasi dapat menerima input nilai, melakukan validasi, menghitung nilai akhir, menentukan predikat, menentukan status kelulusan, dan menampilkan data dalam tabel. Fitur simpan dan hapus data juga dapat digunakan. Aplikasi sudah cukup untuk latihan proyek pemrograman web dasar mahasiswa semester 1.', False)
doc.add_page_break()

heading('BAB V\nPENUTUP', True)
heading('5.1 Kesimpulan')
numbered(['Aplikasi Sistem Perhitungan Nilai Akademik Siswa berhasil dibuat dengan HTML, CSS, dan JavaScript.', 'Aplikasi dapat menghitung nilai akhir dari tugas, UTS, UAS, dan kehadiran secara otomatis.', 'Aplikasi dapat menentukan predikat A sampai E serta status Lulus atau Tidak Lulus.', 'Data nilai dapat disimpan pada localStorage browser.', 'Proyek ini membantu penulis memahami dasar pembuatan aplikasi web dan pengolahan data sederhana.'])
heading('5.2 Saran')
numbered(['Menambahkan fitur edit data siswa.', 'Menambahkan fitur pencarian dan pengurutan data.', 'Menambahkan ekspor data ke Excel atau PDF.', 'Menggunakan database agar data dapat diakses dari perangkat lain.', 'Menambahkan halaman login untuk membedakan pengguna aplikasi.'])

for sec in doc.sections:
    p = sec.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run('Halaman '))
    field = OxmlElement('w:fldSimple')
    field.set(qn('w:instr'), 'PAGE')
    p._p.append(field)
doc.save(OUT)
print(OUT)
