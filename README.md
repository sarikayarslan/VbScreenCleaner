# VB Screen Cleaner

VB5 / VB6 projelerindeki **hiçbir zaman ekranda görünmeyen** form kontrollerini (buton, textbox, label, frame, menü, ActiveX kontroller…) bulan ve bunlardan temizlenmiş bir proje kopyası üreten Windows aracı.

Kaynak projeye **asla yazmaz**: projenin tamamını seçtiğin yeni bir klasöre kopyalar, yalnızca değişen dosyaları yeniden yazar ve yanına ayrıntılı bir rapor bırakır.

---

## İndir

| Sürüm | İndir | Boyut |
|---|---|---|
| **Pencereli uygulama** (önerilen) | ⬇️ [VbScreenCleaner.exe](https://github.com/sarikayarslan/VbScreenCleaner/releases/latest/download/VbScreenCleaner.exe) | ~48 MB |
| **Komut satırı** (toplu işlem / script için) | ⬇️ [VbScreenCleaner.Cli.exe](https://github.com/sarikayarslan/VbScreenCleaner/releases/latest/download/VbScreenCleaner.Cli.exe) | ~35 MB |

Bağlantılar her zaman en son sürümü indirir. Kurulum yok: exe'yi indirip çalıştırman yeterli.

- Windows 10 / 11, 64 bit
- **.NET kurulumu gerekmez** (çalışma ortamı exe'nin içinde; boyutun büyük olmasının nedeni budur)
- Aracı denemek için küçük bir VB5 projesi: [`samples/SampleVb5`](samples/SampleVb5)

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
| Hiç çağrılmayan prosedürleri ölü kod say | Açık | Hiçbir yerden çağrılmayan prosedürlerdeki kod kontrolleri korumaz. Böyle bir prosedür silinen bir kontrolü kullanıyorsa prosedür de silinir (bkz. [Ölü kod](#ölü-kod-hiç-çağrılmayan-prosedürler)). |

---

## Komut satırı

```bat
VbScreenCleaner.Cli.exe analyze "D:\Projeler\EskiProje"
VbScreenCleaner.Cli.exe clean   "D:\Projeler\EskiProje" "D:\Projeler\EskiProje_Temiz"
```

- `analyze`: hiçbir dosyaya dokunmaz, raporu ekrana yazar.
- `clean`: projeyi hedefe kopyalar, temizler, raporu hedefe yazar.
- Ek bayraklar: `--no-load-hidden` (Form_Load kuralını kapatır), `--no-skip-controls` (Controls koleksiyonu kullanan formları da temizler), `--no-dead-code` (ölü kod analizini kapatır).

---

## Nasıl karar veriyor?

### Bir kontrol ne zaman "hiç görünmez" sayılır?

Aşağıdakilerden biri doğruysa **ve** çalışabilen kodun hiçbir yerinde görünür yapılmıyorsa (`.Visible = True`, `.Visible = degisken`, `.Visible = Not x` …). Gizleme olarak `False`, `0`, `vbFalse` ve bunların parantezli hâlleri (`(False)`) kabul edilir:

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
| Kullanıcı etkileşimi olmadan tetiklenen olayı var | `picX_Resize`, `datX_Reposition`, OCX'lerin kendi olayları |
| Değer atayan satır projedeki bir fonksiyonu çağırıyor | `lblX.Caption = SayacArttir()` — satır silinirse yan etki kaybolur |
| Başka bir formun / modülün kodu bu formun `Controls` koleksiyonunu geziyor | `For Each c In frmX.Controls` → frmX'e hiç dokunulmaz |
| Kontrol dizisinin bazı elemanları görünür | `cmd(0)` gizli, `cmd(1)` görünür |
| Dosya proje klasörünün dışındaki ortak bir modülden kullanılıyor | `..\Ortak\Genel.bas` |

### Kontrolle birlikte neler silinir?

- Tasarımdaki `Begin … End` bloğu (alt kontrolleriyle birlikte)
- Olay prosedürleri: `cmdX_Click`, `cmdX_GotFocus`, `cmdArr_Click(Index As Integer)` …
- **Yalnızca o kontrole değer atayan** ifadeler — gizli bir kontrolün değerinin değişmesinin ekranda karşılığı yoktur:
  - `txtX.Text = "..."`, `txtX = "..."` (varsayılan özellik), `cmdX.Font.Bold = True`, `Set picX.Picture = …`
  - `lstX.Clear`, `lstX.AddItem …`, `lstX.RemoveItem …`, `ctl.Move …`, `ctl.ZOrder`, `ctl.Refresh`, `picX.Cls`
  - `ctl.Visible = False`
  - `If kosul Then txtX.Text = "..."` ve `If kosul Then: txtX.Text = "..."` (tek ifadeli, `Else`'siz tek satır If — satırın tamamı silinir)
  - Diğer formlardan / modüllerden yazılanlar: `frmMain.cmdX.Caption = "..."`

Silme **zincirleme** çalışır: gizli bir butonun `_Click` prosedürü başka bir gizli kontrolü kullanıyorsa, buton silinince o kontrol de bir sonraki turda silinir.

### SSTab sekmeleri

SSTab her sekmedeki kontrolleri `Tab(n).Control(k)= "ad"` satırlarıyla listeler. Sekme üzerindeki gizli bir kontrol silinince:

- `Tab(n).Control(k)` ve `Tab(n).Control(k).Enabled` satırları silinir,
- aynı sekmedeki sonraki kayıtlar yeniden numaralanır (`Control(2)` → `Control(1)` …),
- `Tab(n).ControlCount` güncellenir.

### Ölü kod (hiç çağrılmayan prosedürler)

Hiçbir yerden çağrılmayan bir prosedürün kodu hiçbir zaman çalışmaz. Bu yüzden:

- İçindeki kullanımlar kontrolleri **korumaz** (`Private Sub Eski()` içindeki `s = txtX.Text` gibi).
- Böyle bir prosedür silinen bir kontrolü kullanıyorsa **prosedür de silinir**; onu çağıran diğer ölü prosedürler de (proje derlenmeye devam etsin diye).
- Silinen kontrollerle ilgisi olmayan ölü prosedürlere dokunulmaz.

Başlangıç noktası sayılan (her zaman çalışabilir) prosedürler:

| Prosedür | Neden |
|---|---|
| Olaylar: `Form_Load`, `cmdX_Click`, `Class_Initialize`, `mObj_Olay` (WithEvents), `IArayuz_Metot` (Implements) | VB kendisi çağırır (adında `_` olan her prosedür) |
| Modüldeki `Sub Main` | Başlangıç prosedürü |
| Sınıf, UserControl, PropertyPage, UserDocument'taki `Public` / `Friend` üyeler | Nesne üzerinden dışarıdan çağrılabilir |
| Proje klasörü dışındaki dosyalardaki prosedürler | Başka projeler kullanıyor olabilir |

Bunlardan ulaşılabilen her prosedür de çalışır sayılır. "Ulaşılabilir" geniş tutulur: prosedür adı canlı kodda bir kelime olarak (`Yenile`, `Call Yenile`, `frmX.Yenile`, `AddressOf Yenile`) ya da metin olarak (`CallByName Me, "Yenile", VbMethod`) geçiyorsa çağrılıyor kabul edilir.

Kod yeniden yazılırken:

- `a.Text = "x": Debug.Print "y"` → `Debug.Print "y"` (aynı satırdaki diğer ifadeler korunur)
- `ErrH: txtX.Text = ""` → `ErrH:` ve `10 txtX.Text = ""` → `10` (satır etiketi / satır numarası korunur)
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
- **SİLİNEN KULLANILMAYAN PROSEDÜRLER** — ölü kod olup silinen kontrolleri kullandığı için silinenler
- **GÜNCELLENEN SEKME LİSTELERİ** — SSTab'da kaldırılan `Tab(n).Control(k)` kayıtları
- **SİLİNEN KOD İFADELERİ** — orijinal satır numarasıyla
- **GİZLİ AMA KORUNAN KONTROLLER** — her biri için koruma nedeni ve nedenin geçtiği `dosya:satır`
- **UYARILAR** — klasör dışındaki dosyalar, `Controls` koleksiyonu kullanan modüller, bulunamayan dosyalar

---

## Bilinen sınırlar

- `Form_Load` gizleme işini başka bir yardımcı prosedüre bırakıyorsa (`Call EkraniHazirla`) bu kontroller "Form_Load'da gizlenen" sayılmaz, korunur.
- Ölü kod analizi yalnızca taranan dosyalara bakar. Bir formdaki `Public` prosedür klasörde olmayan başka bir projeden çağrılıyorsa bunu göremez; böyle bir durumda bu seçeneği kapatın.
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

Dağıtım dosyalarını (`dist\VbScreenCleaner.exe` ve `dist\VbScreenCleaner.Cli.exe`) üretmek için:

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
tests/VbScreenCleaner.Tests/  xUnit testleri (227 test)
samples/SampleVb5/          Temel senaryoları içeren örnek VB5 projesi
samples/HardVb5/            Zor senaryolar: SSTab, MDI, UserControl, sınıf, Türkçe (1254) adlar, iki .vbp
build/publish.ps1           Dağıtım betiği
```

### Testler

```bat
dotnet test tests\VbScreenCleaner.Tests
```

- **`samples/SampleVb5`** — her formda tam olarak hangi kontrollerin silindiği / korunduğu ve nedenleri
- **`samples/HardVb5`** — gerçek VB5 dosya biçiminde (CRLF, Windows-1254) zorlayıcı senaryolar:
  5 seviye iç içe frame/picturebox, SSTab sekme listelerinin yeniden numaralanması, MDI form, UserControl, `Property Get`'li sınıf,
  ölü kod zincirleri (`OluA` → `OluB`), `CallByName`, `vbFalse` / `(False)`,
  Türkçe karakterli kontrol adları, `If x Then:`, satır numaraları, tek satırlık ve satıra bölünmüş prosedür başlıkları,
  `Forms(0).x` / `Screen.ActiveForm.x` / `With Me`, `#If`, yan etkili fonksiyon çağrıları, `Resize`/`Reposition` olayları,
  modülden başka formun `Controls` döngüsü, aynı adlı kontroller, iki `.vbp`'nin paylaştığı dosyalar ve hiçbir `.vbp`'de olmayan dosya
- **Rastgele üretilen projeler** — 60 farklı rastgele proje (iç içe frame'ler, `Form_Load` / olay / modül içinde okuma, yazma,
  gösterme, gizleme, koşullu ifadeler; modül yarısında çağrılan, yarısında ölü kod). Aracın kararı, kuralları bağımsız olarak
  uygulayan bir referans hesapla birebir karşılaştırılır
- **Ölçek testi** — 150 form × 200 kontrol (30.000 kontrol) analiz + yazma + doğrulama birkaç saniyede biter
- Her senaryoda: temizlenen çıktının yeniden ayrıştırılabilmesi, silinen kontrollere hiçbir referans kalmaması,
  ikinci temizliğin hiçbir şey silmemesi, değişmeyen dosyaların byte byte aynı olması
