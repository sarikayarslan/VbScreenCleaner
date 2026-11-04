# VbOrphanHandlers

VB5/VB6 projelerinde **kontrolü olmayan olay prosedürlerini** bulan Windows masaüstü aracı.

Formdan bir kontrol silindiğinde `cmdEski_Click`, `txtAd_Change` gibi prosedürleri kodda kalabilir. Bu araç bu prosedürleri satır numaralarıyla raporlar. Kontrol dizilerinde (Index'li kontroller) de çalışır.

## İndir

**[VbOrphanHandlers.exe](https://github.com/sarikayarslan/VbScreenCleaner/releases/latest/download/VbOrphanHandlers.exe)**: Windows 10/11 x64 için tek dosya. .NET kurulumu gerekmez.

## Kullanım

1. `VbOrphanHandlers.exe` uygulamasını açın.
2. **Proje** alanına proje klasörünü (**Klasör...**) ya da `.vbp` dosyasını (**.vbp...**) seçin.
3. Formlar ayrı bir `lib` klasöründe de duruyorsa **Lib klasörü** alanına ekleyin (**Ekle...**). Birden fazla klasörü `;` ile ayırabilirsiniz.
4. **Tara** düğmesine basın. Raporu **Raporu kaydet...** ile `.txt` dosyası olarak kaydedebilirsiniz.

Komut satırından da çalıştırılabilir:

```
VbOrphanHandlers.exe "D:\Uygulama\frm" "D:\Uygulama\lib" -o rapor.txt
```

## Rapor bölümleri

| Bölüm | Anlamı |
|---|---|
| 1. Kontrolü olmayan olay prosedürleri | `Nesne_Olay` adlı prosedürün nesnesi formda yok (`Click`, `Change`, `KeyPress`, `GotFocus` ... bilinen olay adları) |
| 2. Index uyuşmazlıkları | Kontrol dizisi olduğu halde prosedürde `Index As Integer` yok, ya da tekil kontrolün prosedürü `Index` alıyor (VB'de derleme hatası verir) |
| 3. Kontrol dizilerinde index kullanımı | Prosedürde `Select Case Index` (`Case 1`, `Case 1, 2`, `Case 3 To 10`) ya da `If/ElseIf Index = n` ile kullanılan index'ler formdakilerle karşılaştırılır. **Kodda olup formda olmayan** index'ler (ölü dallar, satır numarasıyla) ve **formda olup kodda karşılığı olmayan** index'ler listelenir (`Case Else` varsa belirtilir). `Case Is > 5`, `If Index > n` gibi çözülemeyen ifadeler "Elle kontrol" olarak işaretlenir. `Select Case Index + 1` ve `If Index - 1 = 2` gibi kaydırılmış ifadeler çözülür. `mnuX(Index + 1)` / `txtB(Index)` erişimlerinde formda olmayan elemana denk gelen index'ler **çalışma zamanı hatası riski** olarak raporlanır |
| 4. Formda olmayan sabit index erişimleri | Form kodunda `mnuX(15)` gibi formda olmayan bir elemana sabit erişim. `Load mnuX(n)` ile çalışma zamanında eleman eklenen dizilerde not düşülür |
| 5. Bilinmeyen olay adları | `Nesne_Ad` biçimindeki prosedürün nesnesi kontrol değil ve olay adı listede yok. Yardımcı bir fonksiyon olabilir, elle kontrol edin |

Raporlanmayanlar: `Form_` / `MDIForm_` / `UserControl_` / `PropertyPage_` / `UserDocument_` olayları, `WithEvents` değişkenlerinin olayları, `Implements` arayüz metotları ve yorum satırları.

## Doğrudan çağrılar

Olay prosedürü normal bir Sub olarak koddan da çağrılabilir. Bu yüzden tüm `.frm`, `.ctl`, `.bas` ve `.cls` dosyalarında çağrılar aranır:

- **Aynı formdan:** `mnuX_Click 15`, `Call mnuX_Click(15)`, `Me.mnuX_Click 15`. Tek satırlık `If ... Then mnuX_Click 15` de tanınır.
- **Başka formdan / modülden:** `frmAna.mnuX_Click 15`. `Private` prosedürler dışarıdan çağrılamadığı için sayılmaz.
- **Kontrolü olmayan ama çağrılan prosedür:** 1. bölümde `[ÇAĞRILIYOR, ölü değil: dosya:satır]` olarak işaretlenir.
- **Sabit index ile çağrı:** Formda olmayan bir index (ör. 15) ile çağrılıyorsa o index ölü dal sayılmaz, "ÇAĞRIYLA kullanılıyor" olarak gösterilir. Bu index'ler `mnuX(Index)` erişimlerinin çalışma zamanı hatası kontrolüne de katılır.
- **Elle kontrol gerektiren çağrılar:** Değişkenle çağrılar (`mnuX_Click i`), form değişkeni üzerinden çağrılar (`f.mnuX_Click 3`) ve `CallByName(frm, "mnuX_Click", ...)` "Elle kontrol" olarak işaretlenir.

## Notlar

- Taranan dosyalar: `.frm`, `.ctl`, `.pag`, `.dob` (tek bir form dosyası da verilebilir); çağrılar için ayrıca `.bas` ve `.cls`. Bir `.vbp` verildiğinde yalnızca projede listelenen dosyalar (`Form=`, `Module=`, `Class=` ...) taranır. Formlarla modüller ayrı klasördeyse lib klasörünü de ekleyin.
- Alt çizgi içeren kontrol adları (`cmd_Kaydet_Click`) ve satıra bölünmüş (` _`) prosedür başlıkları desteklenir.
- Dosyalar Windows-1254 ile **yalnızca okunur**. Projede hiçbir dosya değiştirilmez.

## Derleme

.NET 9 SDK gerekir.

```
dotnet build src/VbOrphanHandlers -c Release
```

Tek dosya exe oluşturmak için (çalıştırmak için .NET kurulumu gerekmez):

```
dotnet publish src/VbOrphanHandlers -c Release -r win-x64 --self-contained -p:PublishSingleFile=true -o dist
```
