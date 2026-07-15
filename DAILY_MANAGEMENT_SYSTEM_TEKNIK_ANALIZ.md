# DAILY MANAGEMENT SYSTEM(3).xlsm — Teknik Analiz

- Dosya boyutu: 29.9 MB
- Sekme sayısı: 22
- VBA/Makro: Var
- Pivot tablo: 2
- Excel tablosu: 3
- Harici link: 0
- Grafik nesnesi: 0

## Mimari

Dosya üç katmanlıdır: (1) yönetim dashboardları, (2) rapor/çıktı ve ara çalışma sayfaları, (3) NAV/ERP ham veri sekmeleri. Dashboardların çoğu makrolarla yeniden oluşturulur; yalnız SALES bölümünde yoğun dinamik Excel formülleri de vardır.

## MAIN

Ana giriş / başlık sayfası. Güncel tarihi TODAY() ile gösterir ve satış dashboard güncellemesini başlatan UPDATE düğmesi içerir.

- Kullanılan alan: `D4:H4`
- Hücre kaydı: 2
- Değer hücresi: 2
- Formül hücresi: 1
- Benzersiz formül kalıbı: 1
- Birleştirilmiş alan: 0
- Gizli satır: 0
- Öne çıkan fonksiyonlar: TODAY (1)
- Düğmeler/makrolar:
  - UPDATE → UpdateSalesDashboard
- Temsilî formüller:
  - `H4` (1 hücre kalıbı): `TODAY()`

## SALES

Günlük satış yönetim dashboardu. Gün, lokasyon ve belge türü filtreleriyle ciro, maliyet, brüt kâr, marj, teslimat/fatura/müşteri/ürün KPI’ları, geçmiş gün karşılaştırmaları ve top-10 listeleri üretir.

- Kullanılan alan: `A1:AC48`
- Hücre kaydı: 511
- Değer hücresi: 152
- Formül hücresi: 21
- Benzersiz formül kalıbı: 20
- Birleştirilmiş alan: 74
- Gizli satır: 0
- Öne çıkan fonksiyonlar: IF (42), SUMPRODUCT (15), COUNTA (10), LEFT (9), LOWER (9), TRIM (9), ABS (6), IFERROR (6), AVERAGE (2), SUMIFS (1)
- Düğmeler/makrolar:
  - HENT GAMLE OMSÆTNINGER → HentGamleOmsaetninger
  - UPDATE → UpdateSalesDashboard
  - FREEZE → FreezeExcel
  - SAVE AS PDF → GemSalesDashboardSomPDF
  - GET CREDITNOTES → VisKreditnotaFraIgaar
  - VIS... buttons → VisManglendeKunder / VisManglendeVarer / VisNyeKunder / VisNyeVarer
- Temsilî formüller:
  - `O45` (2 hücre kalıbı): `IFERROR(SUMPRODUCT((salesRAW!A21:A999968=$G$4)*(salesRAW!C21:C999968="Salgsleverance")*(IF($C$31="",1,salesRAW!G21:G999968=$C$31))*salesRAW!K21:K999968)/COUNTA(_xlfn.UNIQUE(_xlfn._xlws.FILTER(salesRAW!D21:D999968,(salesRAW!D21:D999968<>"")*(salesRAW!A21:A999968=$G$4)*(salesRAW!C21:C999968="Salgsleverance")*(IF($C$31="",1,salesRAW!G21:G999968=$C$31))))),0)`
  - `I2` (1 hücre kalıbı): `_xlfn.LET( _xlpm.kunder,_xlfn.UNIQUE(_xlfn._xlws.FILTER(salesRAW!N2:N999949,(salesRAW!N2:N999949<>"")*(LEFT(LOWER(TRIM(salesRAW!O2:O999949)),7)<>"kontant"))), _xlpm.navne,_xlfn.XLOOKUP(_xlpm.kunder,salesRAW!N:N,salesRAW!O:O,""), _xlpm.oms,_xlfn.MAP(_xlpm.kunder,_xlfn.LAMBDA(_xlpm.k,SUMIFS(salesRAW!K:K,salesRAW!N:N,_xlpm.k))), _xlfn.TAKE(_xlfn.SORTBY(_xlfn.HSTACK(_xlpm.kunder,_xlpm.navne,_xlpm.oms),_xlpm.oms,-1),10))`
  - `M2` (1 hücre kalıbı): `_xlfn.LET(_xlpm.varer,_xlfn.UNIQUE(_xlfn._xlws.FILTER(salesRAW!E2:E999949,(salesRAW!E2:E999949<>"")*(LEFT(salesRAW!F2:F999949,4)<>"Pant")*(LEFT(salesRAW!F2:F999949,9)<>"Kasse med"))),_xlpm.besk,_xlfn.XLOOKUP(_xlpm.varer,salesRAW!E:E,salesRAW!F:F,""),_xlpm.antal,_xlfn.MAP(_xlpm.varer,_xlfn.LAMBDA(_xlpm.v,SUMPRODUCT((salesRAW!E2:E999949=_xlpm.v)*ABS(salesRAW!H2:H999949)))),_xlfn.TAKE(_xlfn.SORTBY(_xlfn.HSTACK(_xlpm.varer,_xlpm.besk,_xlpm.antal),_xlpm.antal,-1),10))`
  - `AA8` (1 hücre kalıbı): `_xlfn._xlws.SORT(_xlfn.UNIQUE(_xlfn._xlws.FILTER(salesRAW!G2:G999999,salesRAW!G2:G999999<>"")))`
  - `AC8` (1 hücre kalıbı): `_xlfn._xlws.SORT(_xlfn.UNIQUE(_xlfn._xlws.FILTER(salesRAW!C2:C999999,salesRAW!C2:C999999<>"")))`
  - `B33` (1 hücre kalıbı): `SUMPRODUCT((salesRAW!A2:A999999=$G$4)*(IF($C$31="",1,salesRAW!G2:G999999=$C$31))*(IF($H$31="",1,salesRAW!C2:C999999=$H$31))*salesRAW!K2:K999999)`
  - `F33` (1 hücre kalıbı): `SUMPRODUCT((salesRAW!A2:A999999=$G$4)*(IF($C$31="",1,salesRAW!G2:G999999=$C$31))*(IF($H$31="",1,salesRAW!C2:C999999=$H$31))*ABS(salesRAW!L2:L999999))`
  - `J33` (1 hücre kalıbı): `SUMPRODUCT((salesRAW!A21:A999968=$G$4)*(IF($C$31="",1,salesRAW!G21:G999968=$C$31))*(IF($H$31="",1,salesRAW!C21:C999968=$H$31))*salesRAW!K21:K999968)-SUMPRODUCT((salesRAW!A21:A999968=$G$4)*(IF($C$31="",1,salesRAW!G21:G999968=$C$31))*(IF($H$31="",1,salesRAW!C21:C999968=$H$31))*ABS(salesRAW!L21:L999968))`

## salgs u fortjeneste

Kârsız veya sorunlu satış satırlarının çalışma/çıktı alanı. 9.020 satıra kadar satış detayı içerir; mevcut içerik çoğunlukla sonuç değeridir, formül yoktur.

- Kullanılan alan: `A1:K9020`
- Hücre kaydı: 83,297
- Değer hücresi: 81,465
- Formül hücresi: 0
- Benzersiz formül kalıbı: 0
- Birleştirilmiş alan: 0
- Gizli satır: 0

## INVENTORY & SUPPLY CHAIN

Stok ve tedarik kontrol raporu. Ürün, mevcut stok, son 14 gün satışı, maliyet, son alış fiyatı, açık satın alma belgesi/adedi ve en yakın teslimat tarihini birleştirir.

- Kullanılan alan: `B4:J2365`
- Hücre kaydı: 21,209
- Değer hücresi: 17,901
- Formül hücresi: 1
- Benzersiz formül kalıbı: 1
- Birleştirilmiş alan: 0
- Gizli satır: 0
- Öne çıkan fonksiyonlar: TODAY (1)
- Düğmeler/makrolar:
  - GENERATE → GenerateUnsoldProducts
  - GEM SOM PDF → SaveUnsoldListAsPDF
  - GET PURCHASE INFO → AddPurchaseInfoToUnsoldList
- Temsilî formüller:
  - `I4` (1 hücre kalıbı): `TODAY()`

## CASHFLOW

Finansal KPI ekranı. Ciro, COGS, Inventory Days, DSO, DPO ve CCC metriklerini; hedef, durum ve aksiyon yorumlarıyla gösterir.

- Kullanılan alan: `B4:L1000`
- Hücre kaydı: 10,906
- Değer hücresi: 62
- Formül hücresi: 1
- Benzersiz formül kalıbı: 1
- Birleştirilmiş alan: 8
- Gizli satır: 0
- Öne çıkan fonksiyonlar: TODAY (1)
- Düğmeler/makrolar:
  - GENERATE CASHFLOW → GenerateCashflow
  - GEM SOM PDF → GemCashflowSomPDF
- Temsilî formüller:
  - `J4` (1 hücre kalıbı): `TODAY()`

## TAB VIND

Stok aşağı/yukarı düzenlemelerini belge ve ürün bazında eşleştirir; eşleşmeyen nedregulering/opregulering kayıtlarını ve maliyet farklarını gösterir.

- Kullanılan alan: `A4:O388`
- Hücre kaydı: 5,710
- Değer hücresi: 220
- Formül hücresi: 0
- Benzersiz formül kalıbı: 0
- Birleştirilmiş alan: 6
- Gizli satır: 0
- Düğmeler/makrolar:
  - UPDATE → Update_TAB_VIND
  - GEM SOM PDF → Save_TAB_VIND_As_PDF

## SALDORAPPORT

Alacak ve tahsilat takip ekranı. Müşteri bakiyesi, vadesi geçmiş tutar, son fatura, son 14 gün ödeme, 7 günlük saldo değişimi, yeni vade, risk, kredi sorumlusu ve not alanlarını gösterir.

- Kullanılan alan: `B2:AA50000`
- Hücre kaydı: 599,992
- Değer hücresi: 6,202
- Formül hücresi: 2
- Benzersiz formül kalıbı: 2
- Birleştirilmiş alan: 2
- Gizli satır: 0
- Öne çıkan fonksiyonlar: TODAY (1)
- Düğmeler/makrolar:
  - UPDATE → OpdaterSaldoRapport
  - UPDATE KPI's → OpdaterKPIs
  - SAVE AS PDF → GemKPIsSomPDF / GemSaldoRapportSomPDF
- Temsilî formüller:
  - `AA2` (1 hücre kalıbı): `_xlfn._xlws.SORT(_xlfn.UNIQUE(_xlfn._xlws.FILTER(saldoRAW!O:O,(saldoRAW!O:O<>"")*(saldoRAW!O:O<>"Kredit håndtering"))))`
  - `G4` (1 hücre kalıbı): `TODAY()`

## ABC ANALYSIS

Aylık ürün portföy analizi. Ürün bazında fatura/müşteri/adet/ciro/maliyet, delivery–cash&carry kırılımı, marj, birim kâr, delist score/status, kümülatif pay ve ABC sınıfı üretir.

- Kullanılan alan: `A4:T100000`
- Hücre kaydı: 1,999,846
- Değer hücresi: 41,704
- Formül hücresi: 1
- Benzersiz formül kalıbı: 1
- Birleştirilmiş alan: 0
- Gizli satır: 1,042
- Öne çıkan fonksiyonlar: TODAY (1)
- Düğmeler/makrolar:
  - UPDATE → Update_ABC_Analysis
  - SAVE AS PDF → Save_ABC_Analysis_As_PDF
  - GENERATE DELIST → Generate_Delist_Table
- Temsilî formüller:
  - `H4` (1 hücre kalıbı): `TODAY()`

## OPERATIONS

Henüz geliştirilmemiş operasyon dashboard yer tutucusu.

- Kullanılan alan: `D4`
- Hücre kaydı: 1
- Değer hücresi: 1
- Formül hücresi: 0
- Benzersiz formül kalıbı: 0
- Birleştirilmiş alan: 0
- Gizli satır: 0

## STAFF

Henüz geliştirilmemiş personel dashboard yer tutucusu.

- Kullanılan alan: `D4`
- Hücre kaydı: 1
- Değer hücresi: 1
- Formül hücresi: 0
- Benzersiz formül kalıbı: 0
- Birleştirilmiş alan: 0
- Gizli satır: 0

## COUNTING

Cycle counting şablonu. Dört lokasyon için sistem miktarı, depo sayımı ve fark kolonları üretir; çıktı oluşturma makrosu bulunur.

- Kullanılan alan: `B4:S166`
- Hücre kaydı: 2,833
- Değer hücresi: 643
- Formül hücresi: 620
- Benzersiz formül kalıbı: 5
- Birleştirilmiş alan: 7
- Gizli satır: 0
- Öne çıkan fonksiyonlar: IF (16), AND (16)
- Düğmeler/makrolar:
  - GENERATE → GenerateCountingList
- Temsilî formüller:
  - `J14` (604 hücre kalıbı): `[shared formula continuation]`
  - `J12` (4 hücre kalıbı): `IF(AND(H12="",I12=""),"-",H12-I12)`
  - `M12` (4 hücre kalıbı): `IF(AND(K12="",L12=""),"-",K12-L12)`
  - `P12` (4 hücre kalıbı): `IF(AND(N12="",O12=""),"-",N12-O12)`
  - `S12` (4 hücre kalıbı): `IF(AND(Q12="",R12=""),"-",Q12-R12)`

## salesRAW

NAV satış hareketleri ham verisi. SALES dashboardunun ana kaynağıdır; Q kolonu ürün no + açıklama yardımcı alanıdır.

- Kullanılan alan: `A1:Q9401`
- Hücre kaydı: 159,015
- Değer hücresi: 158,968
- Formül hücresi: 8,598
- Benzersiz formül kalıbı: 2
- Birleştirilmiş alan: 0
- Gizli satır: 0
- Düğmeler/makrolar:
  - SAVE RAW FILE → GemSalesRawSomXlsx
- Temsilî formüller:
  - `Q3` (8463 hücre kalıbı): `[shared formula continuation]`
  - `Q2` (135 hücre kalıbı): `E2 & " - " & F2`

## purchaseRAW

Açık satın alma/sipariş satırları ham verisi. Inventory & Supply Chain raporuna sipariş ve beklenen teslimat bilgisi sağlar.

- Kullanılan alan: `A1:K1580`
- Hücre kaydı: 17,380
- Değer hücresi: 17,332
- Formül hücresi: 0
- Benzersiz formül kalıbı: 0
- Birleştirilmiş alan: 0
- Gizli satır: 0

## inventoryRAW

Ürün/stok kartları ham verisi. Stok, birimler, blokaj/webshop durumları, maliyet, satış fiyatı, marj, tedarikçi ve kategori bilgilerini taşır.

- Kullanılan alan: `A1:S2470`
- Hücre kaydı: 46,930
- Değer hücresi: 38,820
- Formül hücresi: 9,876
- Benzersiz formül kalıbı: 2
- Birleştirilmiş alan: 0
- Gizli satır: 0
- Öne çıkan fonksiyonlar: FALSE (6701), TRUE (3175)
- Temsilî formüller:
  - `H2` (6701 hücre kalıbı): `FALSE()`
  - `G2` (3175 hücre kalıbı): `TRUE()`

## cashRAW

Muhasebe hesap planı ve dönem hareket/saldo ham verisi. CASHFLOW metriklerinin kaynağıdır.

- Kullanılan alan: `A1:H376`
- Hücre kaydı: 3,008
- Değer hücresi: 2,678
- Formül hücresi: 0
- Benzersiz formül kalıbı: 0
- Birleştirilmiş alan: 0
- Gizli satır: 0

## tabvindPosterRAW

Stok düzenleme hareketleri ham verisi. TAB VIND raporunun kaynağıdır.

- Kullanılan alan: `A1:O22`
- Hücre kaydı: 330
- Değer hücresi: 268
- Formül hücresi: 0
- Benzersiz formül kalıbı: 0
- Birleştirilmiş alan: 0
- Gizli satır: 0

## abcRAW

Geniş dönem satış hareketleri ham verisi. ABC ANALYSIS için tarihsel ürün satış tabanıdır.

- Kullanılan alan: `A1:P242063`
- Hücre kaydı: 3,873,008
- Değer hücresi: 3,871,644
- Formül hücresi: 0
- Benzersiz formül kalıbı: 0
- Birleştirilmiş alan: 0
- Gizli satır: 0

## saldoRAW

Müşteri kartı ve güncel bakiye ham verisi. SALDORAPPORT ana müşteri kaynağıdır.

- Kullanılan alan: `A1:R2789`
- Hücre kaydı: 50,202
- Değer hücresi: 41,797
- Formül hücresi: 0
- Benzersiz formül kalıbı: 0
- Birleştirilmiş alan: 0
- Gizli satır: 0
- Düğmeler/makrolar:
  - SAVE RAW FILE → GemSaldoRAWSomXLSX

## saldoPosterRAW

Müşteri hareketleri, ödeme, kalan tutar ve vade ham verisi. SALDORAPPORT ödeme/değişim KPI’larını besler.

- Kullanılan alan: `A1:L12389`
- Hücre kaydı: 148,668
- Değer hücresi: 132,344
- Formül hücresi: 0
- Benzersiz formül kalıbı: 0
- Birleştirilmiş alan: 0
- Gizli satır: 0

## operationsRAW

Boş ham veri yer tutucusu.

- Kullanılan alan: `A1`
- Hücre kaydı: 0
- Değer hücresi: 0
- Formül hücresi: 0
- Benzersiz formül kalıbı: 0
- Birleştirilmiş alan: 0
- Gizli satır: 0

## barcodesRAW

Ürün–barkod–birim eşleştirme veri tabanı; sayım ve ürün tanımlama süreçlerinde kullanılmak üzere hazırlanmış.

- Kullanılan alan: `A1:F8075`
- Hücre kaydı: 48,450
- Değer hücresi: 34,226
- Formül hücresi: 0
- Benzersiz formül kalıbı: 0
- Birleştirilmiş alan: 0
- Gizli satır: 0

## Counting 25-06-2026

COUNTING şablonundan oluşturulmuş tarihli sayım raporu örneği/çıktısı.

- Kullanılan alan: `A1:O10`
- Hücre kaydı: 120
- Değer hücresi: 53
- Formül hücresi: 16
- Benzersiz formül kalıbı: 5
- Birleştirilmiş alan: 6
- Gizli satır: 0
- Öne çıkan fonksiyonlar: IF (4), AND (4)
- Düğmeler/makrolar:
  - Save Counting → SaveCountingAsPDF
- Temsilî formüller:
  - `F8` (12 hücre kalıbı): `[shared formula continuation]`
  - `F7` (1 hücre kalıbı): `IF(AND(D7="",E7=""),"-",E7-D7)`
  - `I7` (1 hücre kalıbı): `IF(AND(G7="",H7=""),"-",H7-G7)`
  - `L7` (1 hücre kalıbı): `IF(AND(J7="",K7=""),"-",K7-J7)`
  - `O7` (1 hücre kalıbı): `IF(AND(M7="",N7=""),"-",N7-M7)`

## Kritik teknik bulgular

1. **Performans riski:** SALES formülleri 999.000+ satırlık tam kolon benzeri aralıkları SUMPRODUCT, FILTER, UNIQUE, MAP ve XLOOKUP ile tekrar tekrar tarıyor. Bu yapı yeniden hesaplamayı ağırlaştırır.
2. **Aşırı geniş tablo aralıkları:** purchaseRAW, inventoryRAW ve cashRAW tabloları 1.048.576. satıra kadar tanımlanmış. Gerçek veri birkaç bin satır olsa da tablo/filtre metadata yükünü büyütür.
3. **ABC ANALYSIS şişkinliği:** Sayfa 100.000 satıra ve yaklaşık 2 milyon hücre kaydına biçim/formül alanı taşırken gerçek rapor yaklaşık 2.093 satırdır. 1.042 gizli satır vardır.
4. **abcRAW ana boyut kaynağı:** 242.063 satır ve yaklaşık 3,87 milyon değer hücresiyle dosyanın en büyük veri sekmesidir.
5. **Dinamik dizi bağımlılığı:** LET, FILTER, UNIQUE, SORTBY, TAKE, HSTACK, MAP, LAMBDA ve XLOOKUP kullanımı nedeniyle dosya güncel Microsoft 365 Excel gerektirir. Eski Excel sürümlerinde `_xlfn`/`#NAME?` riski vardır.
6. **Makro bağımlılığı:** Dashboardların çoğu formülle değil VBA ile üretiliyor. Makrolar devre dışıysa güncelleme, PDF kaydetme, sayım oluşturma ve RAW dışa aktarma işlevleri çalışmaz.
7. **Sabit dönem metinleri:** CASHFLOW ve ABC ANALYSIS üzerinde dönem metni 01.06.2026–30.06.2026 olarak kayıtlı. Makronun bunu güncelleyip güncellemediği VBA kaynak kodu tam decompile edilmeden kesinleştirilemez.
8. **Formula cache yaklaşımı:** Birçok dashboard hücresi formül değil, makro tarafından hesaplanıp sabit değer olarak yazılmış görünüyor. Bu nedenle hücre formül sayısının düşük olması, sayfanın statik olduğu anlamına gelmez.
9. **SALDORAPPORT gizli yardımcı alanı:** AA kolonundaki dinamik liste kredi yönetimi filtresini besliyor; R8:S8 dropdown kaynağı bu yardımcı listeden geliyor.
10. **Pivot bağımlılığı:** SALES sonuçlarının bazı ara analizleri salesRAW tabanlı iki pivot tabloya dayanıyor; cache 9.401 kayıtla en son Rasim Beytula tarafından yenilenmiş.

## Güven sınırı

XML, formül, pivot, kontrol, çizim ve makro atama seviyesinde dosya haritalandı. VBA proje dosyası mevcut ve makro giriş adları tespit edildi; ancak bu ortamda VBA modüllerinin tam kaynak kodu güvenilir biçimde decompile edilemedi. Bu nedenle makroların hücre hücre algoritması, dosya yolları ve hata yönetimi yalnızca buton atamalarından ve üretilmiş sonuçlardan çıkarılabilir.