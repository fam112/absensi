import tkinter as tk
from tkinter import ttk, simpledialog, messagebox
import cv2
import sqlite3
from datetime import datetime
from pyzbar.pyzbar import decode
import csv
from tkinter import filedialog
from openpyxl import workbook
from tkcalendar import dateentry 
import winsound



def init_db():
    conn = sqlite3.connect("absensi.db")
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS absensi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nama TEXT,
            metode TEXT,
            waktu TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    conn.commit()
    conn.close()

def simpan_absensi(nama, metode):
    conn = sqlite3.connect("absensi.db")
    c = conn.cursor()
    waktu = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute("INSERT INTO absensi (nama, metode, waktu) VALUES (?, ?, ?)",
    (nama,metode, waktu))

    conn.commit()
    conn.close()
    print(f"[+] {nama} berhasil ditambahkan {metode} pada waktu {waktu}")


def play_sound():
    winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)

def ambil_data():
    conn = sqlite3.connect("absensi.db")
    c = conn.cursor()
    c.execute("SELECT * FROM absensi ORDER BY ID DESC")
    data = c.fetchall()
    conn.close()
    return data

def ambil_recap():
    conn = sqlite3.connect("absensi.db")
    c = conn.cursor()
    c.execute("""
          SELECT nama,
              SUM(CASE WHEN metode='wajah' THEN 1 ELSE 0 END) as total_wajah,
              SUM(CASE WHEN metode='QR' THEN 1 ELSE 0 END) as total_qr,
              COUNT(*) as total_hadir
            FROM absensi
            GROUP BY nama
            ORDER BY total_hadir DESC
""")
    data = c.fetchall()
    conn.close()
    return data

def deteksi_wajah():
    nama = simpledialog.askstring("Input nama", "masukan nama anda: ")
    if not nama:
        messagebox.showwarning("Peringatan", "nama harus diisi!")
        return

    cap = cv2.VideoCapture(0)
    face_decode = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')
    wajah_terdeteksi = False

    while True:
        success, frame = cap.read()
        if not success:
            break
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        faces = face_decode.detectMultiScale(gray, 1.3, 5)

        for (x, y, w, h) in faces:
            cv2.rectangle(frame, (x, y), (x+w, y+h), (255, 0, 0), 2)
            cv2.putText(frame, "Wajah Terdeteksi", (x, y-10),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.9, (255, 0, 0), 2)
            
            if not wajah_terdeteksi:
                simpan_absensi(nama, "wajah")
                messagebox.showinfo("Berhasil", f"Absensi wajah untuk {nama} berhasil!")
                wajah_terdeteksi = True


        cv2.imshow("Absensi Wajah", frame) 
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
            
    cap.release()
    cv2.destroyAllWindows()

def tampilkan_recap():
    recap_window = tk.Toplevel(root)
    recap_window.title("Recap Absensi")
    recap_window.geometry("600x300")

    cols = ("NAMA", "TOTAL WAJAH", "TOTAL QR", "TOTAL HADIR")
    tree = ttk.Treeview(recap_window, columns=cols, show="headings")
    for col in cols:
        tree.heading(col, text=col)
        tree.column(col, width=140)
    
    rows = ambil_recap()
    for row in rows:
        tree.insert("", tk.END, values=row)
        
    tree.pack(expand=True, fill="both")

def deteksi_QR():
    cap = cv2.VideoCapture(0)
    
    while True:
        success, frame = cap.read()
        if not success:
            break

        for barcode in decode(frame):
            data = barcode.data.decode('utf-8')
            nama = data.strip()

            cv2.putText(frame, f"QR: {nama}", (50, 50),
                        cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)
            x, y, w, h = barcode.rect
            cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)
            
            simpan_absensi(nama, "QR")
            messagebox.showinfo("Berhasil", f"Absensi QR untuk {nama} berhasil!")
            cap.release()
            cv2.destroyAllWindows()
            return  # keluar setelah QR terbaca

        cv2.imshow("Absensi QR", frame)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
    
    cap.release()
    cv2.destroyAllWindows()


def tampilkan_data():
    data_wondow = tk.Toplevel(root)
    data_wondow.title("data absensi")
    data_wondow.geometry("500x300")

    cols = ("ID", "NAMA", "METODE", "WAKTU")
    tree = ttk.Treeview(data_wondow, columns=cols, show="Headings")
    for col in cols:
        tree.heading(col, text=col)
        tree.column(col, width=100)
    
    rows = ambil_data()
    for row in rows:
      tree.insert("", tk.END, values=row)
      
    tree.pack(expand=True, fill="both")

def export_csv():
    data = ambil_data()

    if not data :
        messagebox.showwarning("kosong", "Anda belum ada data absensi untuk diexport!")
        return
    
    file = filedialog.asksaveasfile(
        defaultextension=".cvs",
        filetypes=[("CVS files", "*.csv")],
        title="simpan data absensi"
    )

    if file:
        with open(open, mode="w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)

            writer.writerow({"ID", "NAMA", "METODE", "WAKTU"})

            writer.writerows(data)

        messagebox.showinfo("Berhasil", f"Data absensi berhasil di simpan di {file}")

def export_excel():
    data = ambil_data()

    if not data:
        messagebox.showwarning("kosong", "belum ada data absensi untuk diexport!")
        return
    
    file = filedialog.askopenfilename(
        defaultextension=".x1sx",
        filetypes=[("Exsel files", "*x1sx")],
        title="Simpan Data Absensi"

    )

    if file:
        wb = workbook()
        ws = wb.active
        ws.title = "Absensi"

        for row in data:
            ws.append(row)

        wb.save(file)
        messagebox.showinfo("Berhasil", f"Data absensi berhasil di simpan di {file}")

def filter_absensi():
    filter_window = tk. Toplevel(root)
    filter_window.title("Filter Absensi")
    filter_window.geometry("400x300")

    tk.Label(filter_window, text="Pilih tanggal:").pack(padx=5)
    cal = dateentry(filter_window, width=12, background='drakblue',
                    foreground='white', borderwidth=2, year=2025)
    
    cal.pack(pady=5)

    def tampilkan():
        tanggal = cal.get_date().strtime("%Y-%m-%d")
        conn = sqlite3.connect("absensi.db")
        c = conn.cursor()
        c.execute("SELECT * FROM absensi WHERE DATE(waktu)=?", (tanggal,))
        conn.close()

        tree = ttk.Treeview(filter_window, columns=("ID", "NAMA", "METODE","WAKTU"))
        for col in ("ID", "NAMA", "METODE", "WAKTU"):
            tree.heading(col, text=col)
            tree.column(col, width=100)
        for row in row:
            tree.insert("", tk.END, values=row)
            tree.pack(expand=True, fill="both")

        tk,bool(filter_window, text="tampilkan", command=tampilkan).pack(pady=5)



# GUI Tkinter
root = tk.Tk()
root.title("Aplikasi Absensi")        
root.geometry("300x200")


label = tk.Label(root, text="Selamat datang di aplikasi absensi",
                 font=("Arial", 12))
label.pack(pady=10)  

btn_wajah = tk.Button(root, text="Absensi dengan Wajah", command=deteksi_wajah, width=25)
btn_wajah.pack(pady=5)

btn_qr = tk.Button(root, text="Absensi dengan QR", command=deteksi_QR, width=25)
btn_qr.pack(pady=5)

btn_data = tk.Button(root, text="Lihat data absensi", command=tampilkan_recap, width=25)
btn_data.pack(pady=5)

btn_recap = tk.Button(root, text="Lihat Tampilan Recap", command=tampilkan_recap, width=25)
btn_recap.pack(pady=5)


btn_export = tk.Button(root, text="Export ke CSV", command=export_csv, width=25)
btn_export.pack(pady=5)

btn_export_excel = tk.Button(root, text="Export ke Excel", command=export_excel, width=25)
btn_export_excel.pack(padx=5)

btn_filter = tk.Button(root, text="Filter Absensi", command=filter_absensi, width=25)
btn_filter.pack(pady=5)

btn_exit = tk.Button(root, text="Keluar", command=root.quit, width=25,)
btn_exit.pack(pady=5)

init_db()

root.mainloop()
