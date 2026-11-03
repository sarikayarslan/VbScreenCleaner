# VB Screen Cleaner

VB5 / VB6 projelerindeki **hiçbir zaman ekranda görünmeyen** form kontrollerini (buton, textbox, label, frame, menü, ActiveX kontroller…) bulan ve bunlardan temizlenmiş bir proje kopyası üreten Windows aracı.

Kaynak projeye **asla yazmaz**: projenin tamamını seçtiğin yeni bir klasöre kopyalar, yalnızca değişen dosyaları yeniden yazar ve yanına ayrıntılı bir rapor bırakır.

---

## İndir

### ⬇️ [VbScreenCleaner-win-x64.zip](https://github.com/sarikayarslan/VbScreenCleaner/releases/latest/download/VbScreenCleaner-win-x64.zip)

Bu bağlantı her zaman en son sürümü indirir. Zip'i bir klasöre açıp `VbScreenCleaner.exe`'yi çalıştırman yeterli.

- Windows 10 / 11, 64 bit
- **.NET kurulumu gerekmez** (çalışma ortamı exe'nin içinde)
- Zip içeriği:

| Dosya | Açıklama |
|---|---|
| `VbScreenCleaner.exe` | Pencereli uygulama (önerilen) |
| `VbScreenCleaner.Cli.exe` | Komut satırı sürümü (toplu işlem / script için) |
| `OrnekProje\` | Aracı denemek için küçük bir VB5 projesi |
| `README.md` | Bu dosya |

> Windows SmartScreen "tanınmayan uygulama" uyarısı verirse **Ek bilgi → Yine de çalıştır** deyin (exe imzalı değildir).

---

## Ne işe yarar?

Uzun yıllar geliştirilmiş VB5 projelerinde formlar zamanla şu tür kalıntılarla dolar:

- Properties penceresinde `Visible = False` yapılıp unutulmuş butonlar, textbox'lar, frame'ler
- `Form_Load` içinde `xxx.Visible = False` ile kalıcı olarak gizlenmiş kontroller
- Artık gösterilmeyen menüler ve bunların `_Click` prosedürleri
- Gizli frame'lerin içinde kalmış onlarca kontrol

VB Screen Cleaner tüm projeyi tarar, bu kontrollerden **gerçekten hiç görünmeyen ve başka hiçbir yerden kullanılmayanları** belirler ve bunları; olay prosedürleri ve yalnızca onlara değer atayan kod satırlarıyla birlikte temizler. Emin olamadığı her durumda kontrolü **korur** ve nedenini raporda satır numarasıyla yazar.

---

## Hızlı başlangıç (pencereli uygulama)

1. `VbScreenCleaner.exe`'yi aç.
2. **Kaynak proje klasörü**: `.vbp` dosyanın bulunduğu klasörü seç. (Alt klasörlerdeki tüm `.vbp`'ler de okunur.)
3. **Hedef klasör**: boş ya da var olmayan bir klasör. Boş bırakırsan `<proje>_Temiz` önerilir.
4. **Analiz Et**'e bas.
5. Soldaki ağaçta her form için:
   - ☑ **işaretli satırlar** silinecek kontrollerdir. Silinmesini istemediğin kontrolün işaretini kaldır (içindekilerle birlikte korunur).
   - **Gri `KORUNUYOR:` satırları** gizli ama kullanıldığı için korunan kontrollerdir; neden korunduğu yazar.
   - Bir satıra tıklayınca sağda ayrıntılar görünür: hangi alt kontroller, hangi olay prosedürleri silinecek, tasarım özellikleri, koruma nedenleri.
6. **Temizle ve Kaydet**'e bas. Temiz proje hedef klasöre yazılır, yanında `VbScreenCleaner_Rapor.txt` oluşur.
7. Temizlenen projeyi **VB5 IDE'de açıp derle** ve raporu gözden geçir.

### Seçenekler

| Seçenek | Varsayılan | Açıklama |
|---|---|---|
| Form_Load içinde koşulsuz gizlenen kontrolleri de temizle | Açık | Tasarımda görünür olup `Form_Load` içinde hiçbir koşula bağlı olmadan `Visible = False` yapılan kontrolleri de siler. Kapatırsan yalnızca tasarımda `Visible = False` olanlar değerlendirilir. |
| Controls koleksiyonu kullanan formlara dokunma | Açık | Kodu `Me.Controls(...)` veya `For Each c In Controls` kullanan formlarda hiçbir şey silinmez, çünkü bu formlarda kontrollere çalışma anında üretilen adlarla erişilebilir. **Açık bırakılması önerilir.** |

---

## Komut satırı

```bat
VbScreenCleaner.Cli.exe analyze "D:\Projeler\EskiProje"
VbScreenCleaner.Cli.exe clean   "D:\Projeler\EskiProje" "D:\Projeler\EskiProje_Temiz"
```

- `analyze`: hiçbir dosyaya dokunmaz, raporu ekrana yazar.
- `clean`: projeyi hedefe kopyalar, temizler, raporu hedefe yazar.
- Ek bayraklar: `--no-load-hidden` (Form_Load kuralını kapatır), `--no-skip-controls` (Controls koleksiyonu kullanan formları da temizler).

---

## Nasıl karar veriyor?

### Bir kontrol ne zaman "hiç görünmez" sayılır?

Aşağıdakilerden biri doğruysa **ve** kodun hiçbir yerinde görünür yapılmıyorsa (`.Visible = True`, `.Visible = degisken`, `.Visible = Not x` …):

1. Tasarımda `Visible = False`.
2. Tasarımda görünür ama `Form_Load` / `Form_Initialize` / `MDIForm_Load` içinde **koşulsuz** gizleniyor:
   - `If`, `Select Case`, döngü veya `With` bloğunun içinde **değil**,
   - öncesinde `Exit Sub`, `GoTo`, `Return`, `End` gibi akışı kesen bir satır **yok**,
   - başka bir formun `Form_Load`'u tarafından değil, kendi formunun kodunda.
3. Hiç görünmeyen bir **Frame / PictureBox** (veya başka kapsayıcı) içinde duruyor.

Gizli bir menünün alt menüleri gizli **sayılmaz**: bu yapı genelde `PopupMenu` ile açılan sağ tık menüsüdür.

### Frame içindeki kontroller

- Gizli frame ve içindekiler kullanılmıyorsa → frame içindekilerle birlikte tek parça silinir.
- Gizli frame'in içindeki bir kontrol kullanılıyorsa → frame korunur, kullanılmayan diğer içerik tek tek silinir.
- Gizli frame içindeki bir kontrol koddan `Visible = True` yapılsa bile frame görünmediği için o da silinir.

### Bir kontrol ne zaman korunur?

Gizli olsa bile aşağıdakilerden biri varsa **silinmez**:

| Durum | Örnek |
|---|---|
| Değeri okunuyor | `s = txtData.Text`, `If txtX.Text = "" Then` |
| Parametre olarak gönderiliyor / değişkene atanıyor | `Kaydet txtX`, `Set c = txtX` |
| `With` bloğunda kullanılıyor | `With cmdX` |
| Güvenli olmayan metot | `txtX.SetFocus`, `Load cmd(1)`, `Unload` |
| Popup menü olarak açılıyor | `PopupMenu mnuX` |
| Olay prosedürü koddan çağrılıyor | `cmdX_Click` |
| Adı metin olarak geçiyor | `Controls("txtX")`, `CallByName …, "txtX"` |
| Hangi forma ait olduğu belirsiz erişim | `f.cmdX` (f bir parametre), `With frm` içinde `.cmdX` |
| Başka bir kontrolün tasarım özelliği ona bağlı | UpDown `BuddyControl`, ListView `Icons` / `ImageList`, `DataSource` |
| Klavyeyle tetiklenebilir | `Default` / `Cancel` = True buton, kısayol tuşu tanımlı menü |
| Değeri koddan değişiyor **ve** olay prosedürü var | `chkX.Value = 1` + `chkX_Click` (olay tetiklenir) |
| Kontrol dizisinin bazı elemanları görünür | `cmd(0)` gizli, `cmd(1)` görünür |
| Dosya proje klasörünün dışındaki ortak bir modülden kullanılıyor | `..\Ortak\Genel.bas` |

### Kontrolle birlikte neler silinir?

- Tasarımdaki `Begin … End` bloğu (alt kontrolleriyle birlikte)
- Olay prosedürleri: `cmdX_Click`, `cmdX_GotFocus`, `cmdArr_Click(Index As Integer)` …
- **Yalnızca o kontrole değer atayan** ifadeler — gizli bir kontrolün değerinin değişmesinin ekranda karşılığı yoktur:
  - `txtX.Text = "..."`, `txtX = "..."` (varsayılan özellik), `cmdX.Font.Bold = True`, `Set picX.Picture = …`
  - `lstX.Clear`, `lstX.AddItem …`, `lstX.RemoveItem …`, `ctl.Move …`, `ctl.ZOrder`, `ctl.Refresh`, `picX.Cls`
  - `ctl.Visible = False`
  - `If kosul Then txtX.Text = "..."` (tek ifadeli, `Else`'siz tek satır If)
  - Diğer formlardan / modüllerden yazılanlar: `frmMain.cmdX.Caption = "..."`

Silme **zincirleme** çalışır: gizli bir butonun `_Click` prosedürü başka bir gizli kontrolü kullanıyorsa, buton silinince o kontrol de bir sonraki turda silinir.

Kod yeniden yazılırken:

- `a.Text = "x": Debug.Print "y"` → `Debug.Print "y"` (aynı satırdaki diğer ifadeler korunur)
- `ErrH: txtX.Text = ""` → `ErrH:` (satır etiketi korunur)
- ` _` ile birden fazla satıra bölünmüş ifadeler bütün olarak silinir
- Değişmeyen satırların byte'larına dokunulmaz

### Aynı isimli kontroller farklı formlarda

Her kontrol kendi formuna bağlı tutulur; arama metin olarak yapılmaz:

| Kodda | Nerede | Hangi kontrol |
|---|---|---|
| `cmdSil`, `Me.cmdSil`, `Me!cmdSil` | frmA'nın kodu | yalnızca frmA.cmdSil |
| `frmB.cmdSil` | herhangi bir dosya | yalnızca frmB.cmdSil |
| `cmdSil` | `.bas` / `.cls` | hiçbiri (modül bir formun kontrolüne adıyla erişemez) |
| `f.cmdSil`, `"cmdSil"` | herhangi bir dosya | belirsiz → **tüm formlardaki** cmdSil korunur |

---

## Hangi dosyalar taranır?

- Klasör ve alt klasörlerdeki tüm `.vbp` dosyalarında listelenen `.frm`, `.ctl`, `.pag`, `.dob`, `.dsr`, `.bas`, `.cls` dosyaları.
- Klasörde hiç `.vbp` yoksa klasördeki tüm bu uzantılı dosyalar.
- `.vbp`'nin klasör dışından kullandığı ortak dosyalar (`..\Ortak\x.bas`) **okunur** ama değiştirilmez ve kopyalanmaz. Bu dosyalardan kullanılan kontroller silinmez. Raporda uyarı olarak listelenir; en iyisi kaynak olarak bu dosyaları da içine alan üst klasörü seçmektir.
- Kontroller yalnızca `.frm`, `.ctl`, `.pag`, `.dob` dosyalarından silinir.

### Dosya kodlaması

VB5 dosyaları ANSI'dir (Türkçe için Windows-1254). Araç dosyaları byte düzeyinde korunacak şekilde okuyup yazar: Türkçe karakterler, satır sonları (CRLF) ve dosya sonundaki satır sonu aynen kalır. `.frx` dosyaları olduğu gibi kopyalanır.

---

## Rapor

`VbScreenCleaner_Rapor.txt` her dosya için şunları listeler:

- **SİLİNEN KONTROLLER** — satır numarası ve gizli olma nedeniyle; altında birlikte silinen alt kontroller
- **SİLİNEN OLAY PROSEDÜRLERİ**
- **SİLİNEN KOD İFADELERİ** — orijinal satır numarasıyla
- **GİZLİ AMA KORUNAN KONTROLLER** — her biri için koruma nedeni ve nedenin geçtiği `dosya:satır`
- **UYARILAR** — klasör dışındaki dosyalar, `Controls` koleksiyonu kullanan modüller, bulunamayan dosyalar

---

## Bilinen sınırlar

- `Form_Load` gizleme işini başka bir yardımcı prosedüre bırakıyorsa (`Call EkraniHazirla`) bu kontroller "Form_Load'da gizlenen" sayılmaz, korunur.
- Kullanılmayan hale gelen OCX referansları (`Object = …` satırları) formdan silinmez; zararsızdır.
- Silinen kontrollerin `.frx` içindeki verileri (resimler vb.) dosyada kalır; VB bir sonraki kayıtta temizler.
- Araç VB derleyicisi değildir. Temizlenen projeyi **mutlaka VB5 IDE'de derleyip** ekranları kontrol edin.

---

## Kaynaktan derleme

Gerekenler: [.NET 9 SDK](https://dotnet.microsoft.com/download/dotnet/9.0), Windows.

```bat
git clone https://github.com/sarikayarslan/VbScreenCleaner.git
cd VbScreenCleaner
dotnet build VbScreenCleaner.sln -c Release
dotnet test tests\VbScreenCleaner.Tests
```

Dağıtım paketini (`dist\VbScreenCleaner-win-x64.zip`) üretmek için:

```bat
powershell -ExecutionPolicy Bypass -File build\publish.ps1 -Version 1.0.0
```

### Proje yapısı

```
src/
  VbScreenCleaner.Core/     Ayrıştırıcı ve temizlik mantığı
    VbLexer.cs              VB kodunu token'lara ayırır (string, yorum, Rem, :=, rs!Alan …)
    VbProject.cs            .vbp okuma, dosya yükleme, form tasarımı ve kod ayrıştırma
    Model.cs                Form/kontrol ağacı, prosedürler, kod ifadeleri
    CleanupAnalyzer.cs      Neyin silineceğine karar veren analiz
    CleanupWriter.cs        Kopyalama, dosyaları yeniden yazma, rapor
  VbScreenCleaner/          WinForms arayüzü
  VbScreenCleaner.Cli/      Komut satırı arayüzü
tests/VbScreenCleaner.Tests/  xUnit testleri (105 test)
samples/SampleVb5/          Testlerde kullanılan örnek VB5 projesi
build/publish.ps1           Dağıtım paketi betiği
```

### Testler

Testler `samples/SampleVb5` örnek projesi ve test sırasında üretilen küçük projeler üzerinde çalışır:

- Örnek projede her formda tam olarak hangi kontrollerin silindiği / korunduğu ve nedenleri
- Temizlenen çıktının yeniden ayrıştırılabilmesi, silinen kontrollere hiçbir referans kalmaması, ikinci temizliğin hiçbir şey silmemesi, değişmeyen dosyaların byte byte aynı olması
- Okuma, değer atama ve görünür yapmanın onlarca farklı yazımı
- `Form_Load` içindeki koşullu / koşulsuz gizleme biçimleri
- Windows-1254 Türkçe karakterlerin ve `.frx` içeriğinin korunması
