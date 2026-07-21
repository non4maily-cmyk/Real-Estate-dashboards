"""
وحدة تحميل الملفات من Google Drive باستخدام حساب الخدمة
مطابقة الأسماء تتجاهل فروق الهمزة/الألف تلقائياً (نفس درس "أملاك خاصة")
"""
import io, re, os

def norm_ar(s):
    if s is None: return ''
    s = str(s).strip().replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا')
    return re.sub(r'\s+', ' ', s).strip()

def get_drive_service(key_path):
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
    creds = service_account.Credentials.from_service_account_file(key_path, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds)

def download_by_fuzzy_name(service, folder_id, target_name, dest_dir):
    """
    يدوّر على ملف في المجلد اسمه يطابق target_name بعد تجاهل فروق الهمزة/الألف
    والمسافات الزائدة. يرجع مسار الملف المحمّل محلياً.
    """
    from googleapiclient.http import MediaIoBaseDownload

    target_norm = norm_ar(re.sub(r'\.xlsx$', '', target_name, flags=re.IGNORECASE))
    results = service.files().list(
        q=f"'{folder_id}' in parents and trashed = false",
        fields="files(id, name, modifiedTime)",
        pageSize=100
    ).execute()
    files = results.get('files', [])

    match = None
    for f in files:
        fname_norm = norm_ar(re.sub(r'\.xlsx$', '', f['name'], flags=re.IGNORECASE))
        if fname_norm == target_norm:
            match = f
            break

    if not match:
        available = [f['name'] for f in files]
        raise FileNotFoundError(
            f"لم يُعثر على ملف مطابق لـ '{target_name}' في المجلد.\n"
            f"الملفات الموجودة فعلياً: {available}"
        )

    dest_path = os.path.join(dest_dir, target_name)
    request = service.files().get_media(fileId=match['id'])
    fh = io.FileIO(dest_path, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.close()
    return dest_path, match['modifiedTime']


def download_all_required_files(key_path, folder_id, dest_dir):
    """
    يحمّل الأربع ملفات المطلوبة من المجلد ويرجع dict بمساراتهم المحلية.
    أي ملف مفقود يوقف العملية فوراً (لا تخمين، لا نشر جزئي).
    """
    os.makedirs(dest_dir, exist_ok=True)
    service = get_drive_service(key_path)

    required = {
        'main_file': 'التحصيلات_اليومي.xlsx',
        'advance':   'التحصيلات_المقدمة.xlsx',
        'revenues':  'الايرادات_الشهرية.xlsx',
        'deduct':    'الخصم_المؤقت.xlsx',
    }
    paths = {}
    modified_times = {}
    errors = []
    for key, fname in required.items():
        try:
            path, mtime = download_by_fuzzy_name(service, folder_id, fname, dest_dir)
            paths[key] = path
            modified_times[key] = mtime
        except FileNotFoundError as e:
            errors.append(str(e))

    if errors:
        raise FileNotFoundError("\n".join(errors))

    return paths, modified_times
