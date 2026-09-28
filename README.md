# InMa Display Stock V1

Web app HP untuk membantu karyawan melihat barang yang masih memiliki stok dan mengirim referensinya ke WhatsApp untuk pengecekan display.

## Fungsi V1
- Hanya stok > 0
- Filter merek
- Cari artikel
- Urut stok terbesar
- Foto dari Google Drive (bila service account sudah dikonfigurasi)
- Tombol WhatsApp
- Tombol simpan foto untuk dibagikan ke WhatsApp

## Jalankan lokal
1. Install Python 3.11+
2. `pip install -r requirements.txt`
3. `streamlit run app.py`

## Google Drive
Buat `.streamlit/secrets.toml` dan masukkan service-account Google. Share seluruh folder foto produk ke email service account sebagai Viewer.

Format Secrets:
[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n"
client_email = "..."
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."

## Update stok
Ganti file `stock.xlsx` dengan export stok terbaru dengan struktur kolom yang sama, lalu restart/redeploy aplikasi.
