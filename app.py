import io, re, os, json
from pathlib import Path
import pandas as pd
import streamlit as st
from PIL import Image

st.set_page_config(page_title='InMa Display Stock', page_icon='👟', layout='wide')

BRAND_ALIASES = {
 'NE':'NEW ERA','NEW':'NEW ERA','HYS':'HYS ALINA','ALINA':'HYS ALINA','ALI':'HYS ALINA',
 'DULUX':'DULUX','DLX':'DULUX','DUL':'DULUX','BALANCE':'BALANCE','CQ':'CEQIU','CEQIU':'CEQIU',
 'NAKIS':'NAKIS','GLZ':'GLANZ','GLANZ':'GLANZ','GLAZTON':'GLANZ','GLOWY':'GLOWY','UNIQUE':'UNIQUE',
 'HM':'HONGMING/HM','HONGMING':'HONGMING/HM','FEDOR':'FEDOR','BRE':'BRESLIN','BRESLIN':'BRESLIN',
 'PORTO':'PORTO','POR':'PORTO','VERONTINO':'VERONTINO'
}
FOLDER_IDS = {
 'VERONTINO':'1IJWsL_LJ_xPnCayFmtFxon5PD7NCzAa2','PORTO':'1uB59_lsLuhLmJLR_5cP1Q52TaCZAQKjS','NAKIS':'1V-egUaCn7bvrravsBP9L0PxIVOgx7fW_',
 'DULUX':'1tqb0Lay_tZaznacsnpykU23KdczTGwK_','BALANCE':'1dqeKdCo6ECbez4k_d5X4c7mjObSk-WE8','GLOWY':'16BE3kx-KslNoXCbxnLauy_BXbgcjr-I0',
 'UNIQUE':'13J5WF1NztH2oDp6JJEaBVjQFLAdrks37','HONGMING/HM':'1URq7B3eRnCIompkLjlY9r9mFZh6_ui1q','FEDOR':'1XEdT1ZGvtH6kVlZKzC-qfz_luQ49IY_n',
 'CEQIU':'1ao04qXTIQJDen2Q3lAXn2tUI5uCGIIWL','GLANZ':'1sNXhDMBzwqScCY_ZyzlUEmYG9ezPQtKj','BRESLIN':'1FJgUKVln3NMRwYM57nYUCV4jOcuY91KL',
 'HYS ALINA':'1puGj7l1a5bkHjpwDXz0zGIdWvyl3JJcx','NEW ERA':'1JWwh_iArWxyscUshKEAvLw2ApWL1qaYq'
}

def norm(s): return re.sub(r'[^A-Z0-9]','',str(s).upper())

def parse_brand(name):
    t=str(name).upper().strip().split()
    return BRAND_ALIASES.get(t[0],'LAINNYA') if t else 'LAINNYA'

def article_candidates(name):
    t=str(name).upper().strip().split()
    if not t: return []
    first=t.pop(0)
    if first=='NE' and t and t[0] in {'MB','LB','TG','TK','BB','KC','E','CNC'}: t.pop(0)
    out=[]
    for end in range(len(t),0,-1):
        a=norm(''.join(t[:end]))
        if len(a)>=4: out.append(a)
    return list(dict.fromkeys(out))

@st.cache_data
def load_stock(path):
    raw=pd.read_excel(path, header=None)
    header=None
    for i in range(min(20,len(raw))):
        if str(raw.iloc[i,0]).strip().lower()=='kode barang': header=i; break
    if header is None: raise ValueError('Header Kode Barang tidak ditemukan')
    df=pd.read_excel(path, header=header)
    df.columns=[str(c).strip() for c in df.columns]
    name_col=next(c for c in df.columns if c.lower()=='nama barang')
    stock_col=next(c for c in df.columns if c.lower()=='stock')
    size_col=next((c for c in df.columns if 'ukuran' in c.lower()),None)
    df=df[df[name_col].notna()].copy()
    df['Nama']=df[name_col].astype(str).str.strip()
    df['Stok']=pd.to_numeric(df[stock_col],errors='coerce').fillna(0)
    df=df[df['Stok']>0].copy()
    df['Merek']=df['Nama'].map(parse_brand)
    df['Ukuran']=df[size_col].astype(str) if size_col else ''
    return df

@st.cache_data(ttl=900, show_spinner=False)
def drive_index():
    # Optional Google service-account connection. Without credentials the app still works for stock/search.
    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
        info=dict(st.secrets['gcp_service_account'])
        creds=Credentials.from_service_account_info(info,scopes=['https://www.googleapis.com/auth/drive.readonly'])
        svc=build('drive','v3',credentials=creds,cache_discovery=False)
        rows=[]
        for brand,fid in FOLDER_IDS.items():
            token=None
            while True:
                res=svc.files().list(q=f"'{fid}' in parents and trashed=false",fields='nextPageToken,files(id,name,mimeType)',pageSize=1000,pageToken=token).execute()
                for f in res.get('files',[]):
                    if f.get('mimeType','').startswith('image/'):
                        rows.append({'brand':brand,'id':f['id'],'name':f['name'],'norm':norm(f['name'])})
                token=res.get('nextPageToken')
                if not token: break
        return rows
    except Exception:
        return []

@st.cache_data(ttl=900, show_spinner=False)
def drive_image(file_id):
    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
        info=dict(st.secrets['gcp_service_account'])
        creds=Credentials.from_service_account_info(info,scopes=['https://www.googleapis.com/auth/drive.readonly'])
        svc=build('drive','v3',credentials=creds,cache_discovery=False)
        return svc.files().get_media(fileId=file_id).execute()
    except Exception: return None

def find_photo(name,brand,index):
    pool=[f for f in index if f['brand']==brand]
    for c in article_candidates(name):
        for f in pool:
            if c in f['norm']: return f
    return None

st.title('InMa • Cek Barang Display')
st.caption('Hanya menampilkan barang dengan stok > 0')

stock_path=Path(__file__).with_name('stock.xlsx')
try: df=load_stock(stock_path)
except Exception as e:
    st.error(f'Gagal membaca stock.xlsx: {e}'); st.stop()

idx=drive_index()
brands=['Semua']+sorted([x for x in df['Merek'].unique() if x!='LAINNYA'])+(['LAINNYA'] if 'LAINNYA' in set(df['Merek']) else [])
c1,c2=st.columns([1,1])
with c1: selected=st.selectbox('Merek',brands)
with c2: q=st.text_input('Cari artikel',placeholder='Contoh: 650DLB')

view=df.copy()
if selected!='Semua': view=view[view['Merek']==selected]
if q: view=view[view['Nama'].str.contains(q,case=False,na=False,regex=False)]
view=view.sort_values(['Stok','Nama'],ascending=[False,True])
st.subheader(f'{len(view):,} barang')
if not idx: st.info('Foto Drive belum aktif pada deployment ini. Tambahkan Google Drive service-account di Streamlit Secrets; stok dan pencarian tetap dapat digunakan.')

for _,r in view.iterrows():
    photo=find_photo(r['Nama'],r['Merek'],idx) if idx else None
    with st.container(border=True):
        a,b=st.columns([1,1.35])
        with a:
            imgbytes=drive_image(photo['id']) if photo else None
            if imgbytes:
                try: st.image(Image.open(io.BytesIO(imgbytes)),use_container_width=True)
                except Exception: st.caption('Foto tidak dapat dibuka')
            else: st.markdown('### 📷\n**Foto belum tersedia**')
        with b:
            st.markdown(f"### {r['Nama']}")
            st.markdown(f"**Stok: {int(r['Stok']) if float(r['Stok']).is_integer() else r['Stok']} pasang**")
            if str(r['Ukuran']).strip() not in {'','nan','None'}: st.caption(f"Ukuran: {r['Ukuran']}")
            msg=f"{r['Nama']}\nStok: {r['Stok']:g} pasang\nCek barang ini untuk display."
            import urllib.parse
            st.link_button('Kirim ke WhatsApp', 'https://wa.me/?text='+urllib.parse.quote(msg), use_container_width=True)
            if imgbytes:
                ext='.jpg'
                st.download_button('Simpan foto untuk WhatsApp',data=imgbytes,file_name=norm(r['Nama'])+ext,mime='image/jpeg',use_container_width=True)
