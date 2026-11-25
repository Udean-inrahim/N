from docx import Document
from docx.shared import Cm,Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
O=r'C:\Users\Administrator\Documents\N\Cheatsheet_Sistem_Perhitungan_Nilai_Akademik.docx'
d=Document();s=d.sections[0]
for x in ['top_margin','bottom_margin','left_margin','right_margin']:setattr(s,x,Cm(3))
n=d.styles['Normal'];n.font.name='Times New Roman';n._element.rPr.rFonts.set(qn('w:eastAsia'),'Times New Roman');n.font.size=Pt(12);n.paragraph_format.line_spacing=1.5
def R(r,b=False):r.font.name='Times New Roman';r._element.rPr.rFonts.set(qn('w:eastAsia'),'Times New Roman');r.font.size=Pt(12);r.bold=b
def P(t='',a=None,i=True,b=False):
 p=d.add_paragraph();p.paragraph_format.line_spacing=1.5;p.paragraph_format.space_after=Pt(0)
 if i:p.paragraph_format.first_line_indent=Cm(1.25)
 if a is not None:p.alignment=a
 R(p.add_run(t),b);return p
def H(t,l=1):
 p=P(t,WD_ALIGN_PARAGRAPH.CENTER if l==1 else None,False,True);p.paragraph_format.space_before=Pt(12 if l==1 else 6);p.paragraph_format.keep_with_next=True
def T(h,rs):
 t=d.add_table(rows=1,cols=len(h));t.style='Table Grid';t.alignment=WD_TABLE_ALIGNMENT.CENTER
 for j,x in enumerate(h):
  c=t.rows[0].cells[j];c.text='';R(c.paragraphs[0].add_run(x),True);z=OxmlElement('w:shd');z.set(qn('w:fill'),'D9EAF7');c._tc.get_or_add_tcPr().append(z)
 for row in rs:
  cs=t.add_row().cells
  for j,x in enumerate(row):cs[j].text='';p=cs[j].paragraphs[0];p.paragraph_format.line_spacing=1.5;R(p.add_run(str(x)))
 d.add_paragraph()
P('CHEATSHEET',WD_ALIGN_PARAGRAPH.CENTER,False,True);P('SISTEM PERHITUNGAN NILAI AKADEMIK SISWA',WD_ALIGN_PARAGRAPH.CENTER,False,True);P('Ringkasan penggunaan, rumus, dan fungsi program web.html',WD_ALIGN_PARAGRAPH.CENTER,False)
H('1. TUJUAN APLIKASI',2);P('Aplikasi mengolah nilai akademik siswa. Pengguna memasukkan identitas siswa dan empat komponen nilai. Sistem menghitung nilai akhir, predikat, status kelulusan, lalu menyimpan data di localStorage.')
H('2. CARA MENGGUNAKAN',2);T(['Langkah','Tindakan'],[['1','Buka file web.html memakai peramban.'],['2','Isi Nama Siswa dan NIS.'],['3','Isi nilai Tugas, UTS, UAS, dan Kehadiran (0–100).'],['4','Klik tombol “+ Hitung dan Simpan Nilai”.'],['5','Lihat nilai akhir, predikat, dan status pada tabel.'],['6','Klik Hapus untuk menghapus data siswa; setujui konfirmasi.']])
H('3. INPUT DAN VALIDASI',2);T(['Input','Aturan'],[['Nama Siswa','Wajib diisi.'],['NIS','Wajib diisi dan tidak boleh sama dengan NIS tersimpan.'],['Tugas','Angka 0–100. Bobot 30%.'],['UTS','Angka 0–100. Bobot 30%.'],['UAS','Angka 0–100. Bobot 30%.'],['Kehadiran','Angka 0–100. Bobot 10%.']]);P('Pesan validasi: “Pastikan semua data benar dan nilai berada di antara 0-100!” atau “NIS sudah terdaftar!”.',i=False)
H('4. RUMUS NILAI AKHIR',2);P('Nilai Akhir = (Tugas × 0,30) + (UTS × 0,30) + (UAS × 0,30) + (Kehadiran × 0,10)',WD_ALIGN_PARAGRAPH.CENTER,False,True);T(['Komponen','Bobot'],[['Tugas','30%'],['UTS','30%'],['UAS','30%'],['Kehadiran','10%']])
H('5. PREDIKAT DAN KELULUSAN',2);T(['Nilai Akhir','Predikat','Keterangan','Status'],[['85–100','A','Sangat Baik','Lulus'],['75–84,99','B','Baik','Lulus'],['65–74,99','C','Cukup','Tidak Lulus'],['50–64,99','D','Kurang','Tidak Lulus'],['0–49,99','E','Sangat Kurang','Tidak Lulus']]);P('KKM = 75. Status Lulus jika nilai akhir ≥ 75.',i=False)
d.add_page_break();H('6. CONTOH HITUNG CEPAT',2);T(['Tugas','UTS','UAS','Hadir','Nilai Akhir','Predikat','Status'],[['80','75','90','100','83,50','B','Lulus'],['70','70','70','80','71,00','C','Tidak Lulus']]);P('Contoh pertama: (80 × 0,30) + (75 × 0,30) + (90 × 0,30) + (100 × 0,10) = 83,50.',i=False)
H('7. KARTU STATISTIK',2);T(['Kartu','Arti','Cara Hitung'],[['Jumlah Siswa','Total data siswa tersimpan','dataSiswa.length'],['Rata-rata Kelas','Rata-rata nilai akhir semua siswa','Total nilai akhir ÷ jumlah siswa'],['Jumlah Lulus','Total siswa berstatus Lulus','Filter status === “Lulus”']])
H('8. FUNGSI JAVASCRIPT',2);T(['Fungsi / Konstanta','Kegunaan'],[['STORAGE_KEY','Kunci localStorage: data_nilai_siswa.'],['muatData()','Membaca dan mengubah JSON localStorage menjadi array. Jika gagal, mengembalikan array kosong.'],['simpanData()','Menyimpan dataSiswa ke localStorage dengan JSON.stringify().'],['form submit','Mencegah reload, membaca input, validasi, menghitung nilai, lalu menyimpan data.'],['tampilkanData()','Membuat baris tabel hasil dan memperbarui statistik.'],['escapeHTML(teks)','Mengubah karakter khusus HTML pada nama/NIS agar aman ditampilkan.'],['hapusData(index)','Menampilkan konfirmasi lalu menghapus data sesuai indeks.'],['hitungStatistik()','Menghitung jumlah, rata-rata, dan jumlah siswa lulus.']])
H('9. DATA YANG DISIMPAN',2);P('Setiap siswa disimpan sebagai objek JavaScript berikut:',i=False);P('{ nama, nis, tugas, uts, uas, hadir, akhir, predikat, status }',WD_ALIGN_PARAGRAPH.CENTER,False)
H('10. CATATAN PENTING',2);T(['Kondisi','Keterangan'],[['Data tidak hilang saat refresh','Data tersimpan di localStorage peramban.'],['Data hilang jika','Cache/site data peramban dibersihkan atau aplikasi dibuka dari profil/peramban berbeda.'],['NIS sama','Tidak dapat disimpan dua kali.'],['Tampilan mobile','Pada layar < 650 px, form dan statistik berubah menjadi satu kolom.'],['Hapus data','Tindakan perlu dikonfirmasi sebelum data dihapus.']])
for sec in d.sections:
 p=sec.footer.paragraphs[0];p.alignment=WD_ALIGN_PARAGRAPH.CENTER;R(p.add_run('Halaman '));f=OxmlElement('w:fldSimple');f.set(qn('w:instr'),'PAGE');p._p.append(f)
d.save(O);print(O)
