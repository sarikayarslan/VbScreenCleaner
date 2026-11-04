# VbOrphanHandlers

VB5/VB6 projelerinde **kontrolü olmayan olay prosedürlerini** bulan Windows masaüstü aracı.

Formdan bir kontrol silindiğinde `cmdEski_Click`, `txtAd_Change` gibi prosedürleri kodda kalabilir. Bu araç bu prosedürleri satır numaralarıyla raporlar. Kontrol dizilerinde (Index'li kontroller) de çalışır.

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
| 3. Bilinmeyen olay adları | `Nesne_Ad` biçimindeki prosedürün nesnesi kontrol değil ve olay adı listede yok. Yardımcı bir fonksiyon olabilir, elle kontrol edin |

Raporlanmayanlar: `Form_` / `MDIForm_` / `UserControl_` / `PropertyPage_` / `UserDocument_` olayları, `WithEvents` değişkenlerinin olayları, `Implements` arayüz metotları ve yorum satırları.

- Taranan dosyalar: `.frm`, `.ctl`, `.pag`, `.dob`. Bir `.vbp` verildiğinde yalnızca projede listelenen formlar taranır.
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
