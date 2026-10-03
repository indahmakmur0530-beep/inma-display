import base64
import io
import json
import re
import urllib.parse
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image

st.set_page_config(page_title='InMa Display Stock', page_icon='👟', layout='wide')

SPREADSHEET_ID = '1VgV_6LJ778YBwOUoyZl1yWuqfOmwdEnz7quy1ll-6T4'
PRODUCT_SHEET = 'database produk'
LOG_SHEET = 'Kamera Staff'

BRAND_ALIASES = {
    'NE':'NEW ERA','NEW':'NEW ERA','NEW ERA':'NEW ERA',
    'HYS':'HYS ALINA','ALINA':'HYS ALINA','ALI':'HYS ALINA','HYS ALINA':'HYS ALINA',
    'DULUX':'DULUX','DLX':'DULUX','DUL':'DULUX','BALANCE':'BALANCE',
    'CQ':'CEQIU','CEQIU':'CEQIU','NAKIS':'NAKIS',
    'GLZ':'GLANZ','GLANZ':'GLANZ','GLANZTON':'GLANZ','GLAZTON':'GLANZ',
    'GLOWY':'GLOWY','UNIQUE':'UNIQUE','HM':'HONGMING/HM','HONGMING':'HONGMING/HM','HONGMING/HM':'HONGMING/HM',
    'FEDOR':'FEDOR','BRE':'BRESLIN','BRESLIN':'BRESLIN','PORTO':'PORTO','POR':'PORTO',
    'VEROTINO':'VERONTINO','VERONTINO':'VERONTINO',
    'ATT':'ATT','NEXAS':'NEXAS','HILO\'S':'HILO\'S SAS','HILO\'S SAS':'HILO\'S SAS',
    'UNKNOWN':'UNKNOWN'
}

FOLDER_IDS = {
    'NEW ERA':'1JWwh_iArWxyscUshKEAvLw2ApWL1qaYq',
    'HYS ALINA':'1puGj7l1a5bkHjpwDXz0zGIdWvyl3JJcx',
    'BRESLIN':'1FJgUKVln3NMRwYM57nYUCV4jOcuY91KL',
    'NEXAS':'1dmCTE3EmBDuxjVZUi9Whsft5wHXcFTuC',
    'FEDOR':'1XEdT1ZGvtH6kVlZKzC-qfz_luQ49IY_n',
    'HONGMING/HM':'1URq7B3eRnCIompkLjlY9r9mFZh6_ui1q',
    'UNIQUE':'13J5WF1NztH2oDp6JJEaBVjQFLAdrks37',
    'GLOWY':'16BE3kx-KslNoXCbxnLauy_BXbgcjr-I0',
    'HILO\'S SAS':'1HILpLHiu27ReL-ImQtYYY9SlVTjjjrR_',
    'BALANCE':'1dqeKdCo6ECbez4k_d5X4c7mjObSk-WE8',
    'DULUX':'1tqb0Lay_tZaznacsnpykU23KdczTGwK_',
    'UNKNOWN':'1fjVDpwHeo-x2CychI7X4J4CeFTFBUn2p',
    'NAKIS':'1V-egUaCn7bvrravsBP9L0PxIVOgx7fW_',
    'PORTO':'1uB59_lsLuhLmJLR_5cP1Q52TaCZAQKjS',
    'VERONTINO':'1IJWsL_LJ_xPnCayFmtFxon5PD7NCzAa2',
    'CEQIU':'1ao04qXTIQJDen2Q3lAXn2tUI5uCGIIWL',
    'ATT':'1PU6W_O-89S9Sgm2ZxS9uMl-Bl-GJP3Rb',
    'GLANZ':'1sNXhDMBzwqScCY_ZyzlUEmYG9ezPQtKj'
}


def norm(s):
    return re.sub(r'[^A-Z0-9]', '', str(s or '').upper())


def norm_words(s):
    return re.sub(r'[^A-Z0-9]+', ' ', str(s or '').upper()).strip()


def parse_brand(name):
    t = str(name).upper().strip().split()
    return BRAND_ALIASES.get(t[0], 'LAINNYA') if t else 'LAINNYA'


def canon_brand(value):
    x = norm_words(value)
    return BRAND_ALIASES.get(x, x)


def rupiah(value):
    try:
        return 'Rp {:,.0f}'.format(float(str(value).replace('.', '').replace(',', '.'))).replace(',', '.')
    except Exception:
        return str(value or '-')


def article_candidates(name):
    t = str(name).upper().strip().split()
    if not t:
        return []
    first = t.pop(0)
    if first == 'NE' and t and t[0] in {'MB','LB','TG','TK','BB','KC','E','CNC'}:
        t.pop(0)
    out = []
    for end in range(len(t), 0, -1):
        a = norm(''.join(t[:end]))
        if len(a) >= 4:
            out.append(a)
    return list(dict.fromkeys(out))


@st.cache_data
def load_stock(path, file_version=None):
    raw = pd.read_excel(path, header=None)
    header = None
    for i in range(min(20, len(raw))):
        if str(raw.iloc[i, 0]).strip().lower() == 'kode barang':
            header = i
            break
    if header is None:
        raise ValueError('Header Kode Barang tidak ditemukan')
    df = pd.read_excel(path, header=header)
    df.columns = [str(c).strip() for c in df.columns]
    colmap = {str(c).strip().lower(): c for c in df.columns}
    if 'kode barang' not in colmap or 'nama barang' not in colmap:
        raise ValueError(f"Kolom Kode Barang/Nama Barang tidak ditemukan. Kolom terbaca: {list(df.columns)}")
    code_col = colmap['kode barang']
    name_col = colmap['nama barang']
    stock_col = colmap.get('stok') or colmap.get('stock')
    if stock_col is None:
        raise ValueError(f"Kolom Stok/Stock tidak ditemukan. Kolom terbaca: {list(df.columns)}")
    size_col = next((c for c in df.columns if 'ukuran' in c.lower()), None)
    df = df[df[name_col].notna()].copy()
    df['KodeBarang'] = df[code_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.zfill(7)
    df['Nama'] = df[name_col].astype(str).str.strip()
    df['Stok'] = pd.to_numeric(df[stock_col], errors='coerce').fillna(0)
    df = df[df['Stok'] > 0].copy()
    df['Merek'] = df['Nama'].map(parse_brand)
    df['Ukuran'] = df[size_col].astype(str) if size_col else ''
    df['NamaNorm'] = df['Nama'].map(norm)
    return df


@st.cache_data
def load_prices(path):
    df = pd.read_excel(path, dtype={'Kode Barang': str})
    df.columns = [str(c).strip() for c in df.columns]
    colmap = {str(c).strip().lower(): c for c in df.columns}
    code_col = colmap.get('kode barang')
    price_col = colmap.get('harga jual1') or colmap.get('harga jual 1')
    if code_col is None or price_col is None:
        raise ValueError(f"Kolom Kode Barang/Harga Jual1 tidak ditemukan. Kolom terbaca: {list(df.columns)}")
    out = df[[code_col, price_col]].copy()
    out.columns = ['KodeBarang', 'HargaJual1']
    out['KodeBarang'] = out['KodeBarang'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.zfill(7)
    out['HargaJual1'] = pd.to_numeric(out['HargaJual1'], errors='coerce')
    out = out.dropna(subset=['KodeBarang']).drop_duplicates('KodeBarang', keep='last')
    return out


def share_product_button(image_bytes, filename, message, key):
    """Mobile Web Share: share the actual image file plus text to WhatsApp/other apps."""
    if not image_bytes:
        return
    b64 = base64.b64encode(image_bytes).decode('ascii')
    safe_message = json.dumps(message)
    safe_filename = json.dumps(filename)
    html = f"""
    <button id="share-{key}" style="width:100%;padding:0.65rem 1rem;border:1px solid #d0d0d0;border-radius:0.5rem;background:white;font-size:16px;cursor:pointer;">
      Bagikan foto + harga
    </button>
    <div id="status-{key}" style="font:13px sans-serif;margin-top:6px;color:#666;"></div>
    <script>
    const btn = document.getElementById('share-{key}');
    const status = document.getElementById('status-{key}');
    btn.onclick = async () => {{
      try {{
        const binary = atob('{b64}');
        const bytes = new Uint8Array(binary.length);
        for (let i=0; i<binary.length; i++) bytes[i] = binary.charCodeAt(i);
        const file = new File([bytes], {safe_filename}, {{type:'image/jpeg'}});
        const data = {{text: {safe_message}, files:[file]}};
        if (navigator.share && (!navigator.canShare || navigator.canShare(data))) {{
          await navigator.share(data);
          status.textContent = '';
        }} else {{
          status.textContent = 'Browser ini belum mendukung berbagi foto langsung. Gunakan tombol Simpan foto.';
        }}
      }} catch (e) {{
        if (e.name !== 'AbortError') status.textContent = 'Gagal membuka menu bagikan: ' + e.message;
      }}
    }};
    </script>
    """
    components.html(html, height=62)


@st.cache_resource
def google_clients():
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    info = dict(st.secrets['gcp_service_account'])
    scopes = [
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/spreadsheets',
    ]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    drive = build('drive', 'v3', credentials=creds, cache_discovery=False)
    sheets = build('sheets', 'v4', credentials=creds, cache_discovery=False)
    return drive, sheets


@st.cache_data(ttl=900, show_spinner=False)
def drive_index():
    try:
        drive, _ = google_clients()
        rows = []
        for brand, fid in FOLDER_IDS.items():
            token = None
            while True:
                res = drive.files().list(
                    q=f"'{fid}' in parents and trashed=false",
                    fields='nextPageToken,files(id,name,mimeType)',
                    pageSize=1000,
                    pageToken=token,
                ).execute()
                for f in res.get('files', []):
                    if f.get('mimeType', '').startswith('image/'):
                        rows.append({'brand': brand, 'id': f['id'], 'name': f['name'], 'norm': norm(f['name'])})
                token = res.get('nextPageToken')
                if not token:
                    break
        return rows
    except Exception:
        return []


@st.cache_data(ttl=900, show_spinner=False)
def drive_image(file_id):
    try:
        drive, _ = google_clients()
        return drive.files().get_media(fileId=file_id).execute()
    except Exception:
        return None


def find_photo(name, brand, index):
    pool = [f for f in index if f['brand'] == brand]
    for c in article_candidates(name):
        for f in pool:
            if c in f['norm']:
                return f
    return None


@st.cache_data(ttl=300, show_spinner=False)
def load_products():
    _, sheets = google_clients()
    vals = sheets.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{PRODUCT_SHEET}'!A1:H5000",
    ).execute().get('values', [])
    if not vals:
        return pd.DataFrame(columns=['brand','artikel','harga','ukuran','nama_file','folder','file_id','brand_canon','search_text'])
    header = [norm_words(x).lower().replace(' ', '_') for x in vals[0]]
    rows = []
    for r in vals[1:]:
        r = r + [''] * (len(header) - len(r))
        rows.append(dict(zip(header, r)))
    df = pd.DataFrame(rows)
    for c in ['brand','artikel','harga','ukuran','nama_file','folder','file_id']:
        if c not in df.columns:
            df[c] = ''
    df['brand_canon'] = df['brand'].map(canon_brand)
    df['search_text'] = (df['brand'].astype(str) + ' ' + df['artikel'].astype(str) + ' ' + df['nama_file'].astype(str)).map(norm_words)
    df = df[df['artikel'].astype(str).str.strip() != ''].copy()
    return df


def analyze_photo(image_bytes, selected_brand, candidate_articles):
    from openai import OpenAI
    client = OpenAI(api_key=st.secrets['OPENAI_API_KEY'])
    b64 = base64.b64encode(image_bytes).decode('ascii')
    # Keep prompt compact. The model reads visible codes/text; local ranking checks all articles afterwards.
    shortlist = '\n'.join(candidate_articles[:300])
    prompt = f"""
Anda membantu identifikasi produk sepatu/sandal milik toko InMa.
Merek yang dipilih staf: {selected_brand}.
Perhatikan tulisan pada upper, sol, tag, stiker, kotak, bentuk produk, warna, strap/tali, dan detail desain.
Jangan mengarang kode artikel. Jika kode artikel tidak terlihat, gunakan ciri visual hanya sebagai petunjuk dan turunkan confidence.
Dari daftar artikel toko berikut, sebutkan maksimal 5 artikel yang paling mungkin bila ada bukti yang cukup.

DAFTAR ARTIKEL:
{shortlist}

Kembalikan JSON valid saja:
{{"visible_text":["..."],"category":"...","colors":["..."],"features":["..."],"article_hint":"...","candidate_articles":["..."],"confidence":0}}
confidence 0 sampai 100.
"""
    response = client.responses.create(
        model='gpt-6-luna',
        input=[{
            'role': 'user',
            'content': [
                {'type': 'input_text', 'text': prompt},
                {'type': 'input_image', 'image_url': f'data:image/jpeg;base64,{b64}', 'detail': 'high'},
            ],
        }],
    )
    text = response.output_text.strip()
    text = re.sub(r'^```(?:json)?\s*|\s*```$', '', text, flags=re.I | re.S)
    return json.loads(text)


def rank_candidates(df_brand, ai):
    from rapidfuzz import fuzz
    visible = ' '.join(ai.get('visible_text', []))
    hint = ai.get('article_hint', '')
    ai_candidates = ai.get('candidate_articles', []) or []
    query = norm_words(' '.join([visible, hint] + ai_candidates))
    rows = []
    for idx, r in df_brand.iterrows():
        score = fuzz.WRatio(query, r['search_text']) if query else 0
        art = norm(r['artikel'])
        if any(norm(x) == art for x in ai_candidates):
            score = max(score, 96)
        elif hint and fuzz.ratio(norm(hint), art) >= 80:
            score = max(score, 90)
        rows.append((idx, score))
    rows.sort(key=lambda x: x[1], reverse=True)
    return [(df_brand.loc[i], s) for i, s in rows[:5]]


def stock_match(article, brand, stock_df):
    a = norm(article)
    pool = stock_df[stock_df['Merek'] == brand]
    if not a or pool.empty:
        return None
    exact = pool[pool['NamaNorm'].str.contains(a, regex=False, na=False)]
    if not exact.empty:
        return exact.sort_values('Stok', ascending=False).iloc[0]
    return None


def upload_to_drive(image_bytes, folder_id, filename):
    from googleapiclient.http import MediaIoBaseUpload
    drive, _ = google_clients()
    media = MediaIoBaseUpload(io.BytesIO(image_bytes), mimetype='image/jpeg', resumable=False)
    return drive.files().create(
        body={'name': filename, 'parents': [folder_id]},
        media_body=media,
        fields='id,name,webViewLink',
    ).execute()


def append_log(values):
    _, sheets = google_clients()
    sheets.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{LOG_SHEET}'!A:L",
        valueInputOption='USER_ENTERED',
        insertDataOption='INSERT_ROWS',
        body={'values': [values]},
    ).execute()


def clear_camera_state():
    for k in ['ai_result','ranked','photo_bytes','camera_brand']:
        st.session_state.pop(k, None)


stock_path = Path(__file__).with_name('stock.xlsx')
price_path = Path(__file__).with_name('harga.xlsx')
try:
    stock_version = (stock_path.stat().st_mtime_ns, stock_path.stat().st_size)
    df = load_stock(stock_path, stock_version).copy()
    # Rehitung merek setiap rerun agar perubahan alias/mapping tidak tertahan cache lama.
    df['Merek'] = df['Nama'].map(parse_brand)
    prices = load_prices(price_path)
    df = df.merge(prices, on='KodeBarang', how='left')
except Exception as e:
    st.error(f'Gagal membaca stock.xlsx / harga.xlsx: {e}')
    st.stop()

st.sidebar.title('InMa')
page = st.sidebar.radio('Menu', ['📦 Display Stock', '🔎 Tanya AI', '🎓 Ajarkan Barang'])

if page == '📦 Display Stock':
    st.title('InMa • Cek Barang Display')
    st.caption('Hanya menampilkan barang dengan stok > 0')

    idx = drive_index()
    brands = ['Semua'] + sorted([x for x in df['Merek'].unique() if x != 'LAINNYA']) + (['LAINNYA'] if 'LAINNYA' in set(df['Merek']) else [])
    c1, c2 = st.columns([1, 1])
    with c1:
        selected = st.selectbox('Merek', brands)
    with c2:
        q = st.text_input('Cari artikel', placeholder='Contoh: 650DLB')

    view = df.copy()
    if selected != 'Semua':
        view = view[view['Merek'] == selected]
    if q:
        view = view[view['Nama'].str.contains(q, case=False, na=False, regex=False)]
    view = view.sort_values(['Stok','Nama'], ascending=[False, True])
    st.subheader(f'{len(view):,} barang')
    if not idx:
        st.info('Foto Drive belum aktif pada deployment ini. Periksa Google Drive service-account di Streamlit Secrets.')

    for _, r in view.iterrows():
        photo = find_photo(r['Nama'], r['Merek'], idx) if idx else None
        with st.container(border=True):
            a, b = st.columns([1, 1.35])
            with a:
                imgbytes = drive_image(photo['id']) if photo else None
                if imgbytes:
                    try:
                        st.image(Image.open(io.BytesIO(imgbytes)), width='stretch')
                    except Exception:
                        st.caption('Foto tidak dapat dibuka')
                else:
                    st.markdown('### 📷\n**Foto belum tersedia**')
            with b:
                st.markdown(f"### {r['Nama']}")
                stok_text = int(r['Stok']) if float(r['Stok']).is_integer() else r['Stok']
                st.markdown(f'**Stok: {stok_text} pasang**')
                if str(r['Ukuran']).strip() not in {'','nan','None'}:
                    st.caption(f"Ukuran: {r['Ukuran']}")
                harga = r.get('HargaJual1', '')
                if pd.notna(harga):
                    st.markdown(f"**Harga Jual 1: {rupiah(harga)}**")
                else:
                    st.warning('Harga Jual 1 tidak ditemukan untuk kode barang ini.')
                st.caption(f"Kode Barang: {r['KodeBarang']}")

                # Pesan pelanggan: jangan kirim stok internal. Harga selalu dari Harga Jual1 POS.
                msg = f"{r['Nama']}\nKode: {r['KodeBarang']}\nHarga: {rupiah(harga) if pd.notna(harga) else '-'}"
                if imgbytes and pd.notna(harga):
                    share_product_button(
                        imgbytes,
                        norm(r['Nama']) + '.jpg',
                        msg,
                        key=norm(str(r['KodeBarang']) + r['Nama'])[:40],
                    )
                else:
                    st.link_button('Kirim teks ke WhatsApp', 'https://wa.me/?text=' + urllib.parse.quote(msg), width='stretch')
                if imgbytes:
                    st.download_button('Simpan foto', data=imgbytes, file_name=norm(r['Nama']) + '.jpg', mime='image/jpeg', width='stretch')


elif page == '🔎 Tanya AI':
    st.title('InMa • Tanya AI')
    st.caption('Foto barang → AI mencari kandidat → foto referensi Drive ditampilkan untuk dibandingkan')

    if 'OPENAI_API_KEY' not in st.secrets:
        st.error('OPENAI_API_KEY belum ada di Streamlit Secrets.')
        st.stop()

    try:
        products = load_products()
    except Exception as e:
        st.error(f'Gagal membaca database produk Google Sheets: {e}')
        st.stop()

    available_brands = sorted(set(products['brand_canon'].dropna()) & set(FOLDER_IDS.keys()))
    if not available_brands:
        st.error('Tidak menemukan merek yang cocok antara database produk dan folder Drive.')
        st.stop()

    brand = st.selectbox('Pilih merek', available_brands, key='ask_brand')
    photo = st.camera_input('Foto barang yang ingin dikenali', resolution='720p', key='ask_photo')

    if photo and brand:
        image_bytes = photo.getvalue()
        df_brand = products[products['brand_canon'] == brand].copy()
        article_list = [str(x) for x in df_brand['artikel'].dropna().unique() if str(x).strip()]

        if st.button('Cari kandidat', type='primary', width='stretch'):
            if not article_list:
                st.warning('Artikel untuk merek ini belum ada di database produk.')
            else:
                with st.spinner('AI membaca foto dan mencari kandidat...'):
                    try:
                        ai = analyze_photo(image_bytes, brand, article_list)
                        ranked = rank_candidates(df_brand, ai)
                        st.session_state['ai_result'] = ai
                        st.session_state['ranked'] = [(r.to_dict(), s) for r, s in ranked]
                        st.session_state['photo_bytes'] = image_bytes
                        st.session_state['camera_brand'] = brand
                    except Exception as e:
                        st.error(f'Identifikasi gagal: {e}')

    if st.session_state.get('ranked') and st.session_state.get('camera_brand') == brand:
        ai = st.session_state.get('ai_result', {})
        ranked = st.session_state['ranked']
        st.caption(f"Keyakinan AI: {ai.get('confidence', 0)}%")
        if ai.get('visible_text'):
            st.caption('Tulisan terbaca: ' + ', '.join(ai.get('visible_text', [])))

        idx = drive_index()
        st.subheader('Kandidat yang paling mungkin')

        for nomor, (r, score) in enumerate(ranked, start=1):
            sm = stock_match(r.get('artikel', ''), brand, df)
            nama = sm['Nama'] if sm is not None else str(r.get('artikel', ''))
            harga = sm['HargaJual1'] if sm is not None and pd.notna(sm.get('HargaJual1')) else ''
            stok = sm['Stok'] if sm is not None else None

            # Prefer the product-sheet file_id; otherwise find the matching reference photo in the brand Drive folder.
            ref_id = str(r.get('file_id', '') or '').strip()
            ref_bytes = drive_image(ref_id) if ref_id else None
            if not ref_bytes and idx:
                ref_photo = find_photo(nama, brand, idx)
                if ref_photo:
                    ref_bytes = drive_image(ref_photo['id'])

            with st.container(border=True):
                if ref_bytes:
                    try:
                        st.image(Image.open(io.BytesIO(ref_bytes)), width='stretch')
                    except Exception:
                        st.caption('Foto referensi ditemukan tetapi tidak dapat dibuka.')
                else:
                    st.caption('📷 Foto referensi belum ditemukan di Drive.')

                st.markdown(f'### {nomor}. {nama}')
                st.markdown(f'**Artikel: {r.get("artikel", "-")}**')
                if harga != '':
                    st.markdown(f'**Harga Jual 1: {rupiah(harga)}**')
                else:
                    st.markdown('**Harga Jual 1: -**')
                if stok is not None:
                    st.markdown(f'**Stok: {stok:g} pasang**')
                st.caption(f'Kecocokan kandidat: {score:.0f}%')

        st.info('Bandingkan barang yang difoto dengan foto referensi Drive. Hasil AI adalah kandidat, bukan konfirmasi otomatis.')

else:
    st.title('InMa • Ajarkan Barang')
    st.caption('Foto barang asli → staf memilih artikel yang BENAR → foto disimpan sebagai contoh artikel tersebut')
    st.info('Mode ini tidak menebak artikel. Label berasal dari staf dan disimpan sebagai data contoh.')

    try:
        products = load_products()
    except Exception as e:
        st.error(f'Gagal membaca database produk Google Sheets: {e}')
        st.stop()

    available_brands = sorted(set(products['brand_canon'].dropna()) & set(FOLDER_IDS.keys()))
    if not available_brands:
        st.error('Tidak menemukan merek yang cocok antara database produk dan folder Drive.')
        st.stop()

    staff = st.text_input('Nama staf', placeholder='Contoh: Dodi', key='teach_staff')
    brand = st.selectbox('Pilih merek', available_brands, key='teach_brand')

    df_brand = products[products['brand_canon'] == brand].copy()
    df_brand['artikel'] = df_brand['artikel'].astype(str).str.strip()
    df_brand = df_brand[df_brand['artikel'] != ''].drop_duplicates('artikel')

    search = st.text_input('Cari artikel yang benar', placeholder='Ketik kode artikel', key='teach_search')
    if search.strip():
        s = search.strip()
        filtered = df_brand[
            df_brand['artikel'].str.contains(s, case=False, na=False, regex=False) |
            df_brand['nama_file'].astype(str).str.contains(s, case=False, na=False, regex=False)
        ].copy()
    else:
        filtered = df_brand.copy()

    selected_row = None
    if filtered.empty:
        st.warning('Artikel tidak ditemukan untuk merek ini.')
    else:
        filtered = filtered.sort_values('artikel').head(300)
        labels, by_label = [], {}
        for _, pr in filtered.iterrows():
            artikel = str(pr.get('artikel', '')).strip()
            sm = stock_match(artikel, brand, df)
            nama = sm['Nama'] if sm is not None else artikel
            kode = sm['KodeBarang'] if sm is not None else ''
            harga = sm['HargaJual1'] if sm is not None else ''
            stok = sm['Stok'] if sm is not None else None
            extra = []
            if kode:
                extra.append(f'kode {kode}')
            if pd.notna(harga) and str(harga).strip():
                extra.append(rupiah(harga))
            if stok is not None:
                extra.append(f'stok {stok:g}')
            label = artikel + (f" | {' | '.join(extra)}" if extra else '')
            labels.append(label)
            by_label[label] = (pr.to_dict(), sm, nama, harga)
        choice = st.selectbox('Pilih artikel yang BENAR', labels, key='teach_choice')
        selected_row = by_label.get(choice)

    photo = st.camera_input('Foto barang asli', resolution='720p', key='teach_photo')

    if selected_row is not None:
        pr, sm, nama, harga = selected_row
        with st.container(border=True):
            st.markdown(f'### {nama}')
            st.markdown(f'**Artikel: {pr.get("artikel", "-")}**')
            if sm is not None:
                st.caption(f'Kode Barang POS: {sm["KodeBarang"]}')
                st.markdown(f'**Harga Jual 1: {rupiah(harga)}**')
                st.markdown(f'**Stok: {sm["Stok"]:g} pasang**')

    if photo and selected_row is not None:
        image_bytes = photo.getvalue()
        st.success('Foto siap diberi label. Pastikan artikel di atas benar sebelum menyimpan.')

        if st.button('Simpan sebagai contoh artikel ini', type='primary', width='stretch'):
            if not staff.strip():
                st.warning('Isi nama staf terlebih dahulu.')
            else:
                pr, sm, nama, harga = selected_row
                folder_id = FOLDER_IDS.get(brand)
                if not folder_id:
                    st.error('Folder merek belum dipetakan.')
                else:
                    artikel = str(pr.get('artikel', '')).strip()
                    safe_article = re.sub(r'[^A-Za-z0-9._-]+', '_', artikel or 'BARANG').strip('_')
                    safe_brand = re.sub(r'[^A-Za-z0-9._-]+', '_', brand).strip('_')
                    kode = str(sm['KodeBarang']) if sm is not None else ''
                    safe_code = re.sub(r'[^A-Za-z0-9._-]+', '_', kode).strip('_')
                    parts = [safe_brand, safe_article]
                    if safe_code:
                        parts.append(safe_code)
                    parts.append(datetime.now().strftime("%Y%m%d_%H%M%S"))
                    filename = '_'.join(parts) + '.jpg'
                    try:
                        f = upload_to_drive(image_bytes, folder_id, filename)
                        append_log([
                            datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                            staff.strip(), filename, brand, artikel, nama,
                            harga if pd.notna(harga) else '', '', 'DIAJARKAN',
                            folder_id, f.get('id', ''),
                            f'Label diberikan langsung oleh staf. Kode POS: {kode}'
                        ])
                        drive_index.clear()
                        load_products.clear()
                        st.success(f'Berhasil. Foto tersimpan sebagai contoh artikel {artikel}.')
                        if f.get('webViewLink'):
                            st.link_button('Buka foto contoh di Drive', f['webViewLink'])
                    except Exception as e:
                        st.error(f'Gagal menyimpan contoh ke Drive/Google Sheet: {e}')

