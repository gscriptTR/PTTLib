// Sayfalar
  const sfMain = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Ptt Kargo')
  const sfEtiket = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ETIKET')
  const sfAdres = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Adres Defteri')
  const sfAyar = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AYARLAR')
  const sfLog = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LOG')

//

/** TANITIM
 * Etiket sayfasında gerekli satırlar doldurulur
 * Ya da, adres defterinden seçilen(seciliAdres isimlendirmesi ile) müşteri bilgileri ekrana otomatik getirilir(namedRange)
 * Menüden etiket çıkara basılır
 * Etiket çıkarılan satırlar bir numara ile gruplanır(Z sütunu)
 * Satırlara sırayla kaldığımız yerden başlayarak barkod verilir
 * Etiket sayfasındaki fazlalık satırlar gizlenir
 * Barkod sırasında son kalan sıra kaydedilir.
 * PTT dosyası da grupta gizli gün bilgisi ile otomatik oluşturulur
 * 
*/

//Versiyon Sistemi
function versiyonGoster(){
  /*
  * 220501: Ptt firması şablonuna uygun sayfa tasarlandı
  * 220504: İl ilçe altyapısı normal düzene taşındı
  * 220506: Etiket sayfası düzenlendi
  * 220510: Barkod aralığı belirlendi
  * 220511: Barkod fonksiyonu yazıldı
  * 220512: PTT datası sorgusu oluşturuldu
  * 220513: Adres defteri eklendi
  * 220514: Etiket sayfasında fazlalıklar gizleniyor.
  * 220524: Zorunlu alan kontrolü eklendi
  * 220524: Desi kontrolleri eklendi
  * 220524: ST kontrolü eklendi
  * 220524: Etiket sayfasındaki tek sefer etiket sayısı arttırıldı.
  * 220524: Etiket Girişi fonksiyonu ile başlangıç değerleri otomatik giriliyor.
  * 220602: Kütüphane kullanımına geçti
  * 220602: Saat hatası giderildi(GMT)
  */
  let vers = 220605

  SpreadsheetApp.getActive().toast(vers, "VERSİYON")

  //Versiyonu dönelim
  return vers
}

//Açılış fonksiyonu
function onOpen(){ 
  
  SpreadsheetApp.getUi()
  .createMenu('_OPERASYON')
  .addItem('✨ Etiket Girişi', 'etiketGirisi')
  .addItem('🚀 Adres Çağır >>', 'adresDefterinden')
  .addItem('🏷️ Etiket Çıkar', 'kargoEtiket')
  .addSeparator()
  .addItem('Vers. : ' + versiyonGoster(), 'versiyonGoster')
  .addToUi()

}
//Adres defterinden veri getirmek için

function adresDefterinden(){

  //Seçilen Firma bilgileri
  let arrSecilen = sfAdres.getDataRange().getDisplayValues().find(item => {
    return item[1] == sfMain.getRange("seciliAdres").getValues()
  })

  
  if(arrSecilen == undefined)
    throw("Bilgiler bulunamadı")


  //Seçilen satırı taşıyalım. Bilgi verelim
  sfMain.appendRow(arrSecilen)
  SpreadsheetApp.getActive().toast(arrSecilen[1] + " firması için kargo")
  
  //Etiket girişine geçelim
  etiketGirisi(sfMain.getLastRow())
  
 
}

//Etiket girişindeki temel bilgiler
function etiketGirisi(satirNo = null){

  //İşlem yapılacak satır
  let yeniSatir = satirNo || sfMain.getLastRow()+1
  
  //Yeni Satıra konumlanıp, Varsayılan bilgileri yazalım
  sfMain.getRange("B"+yeniSatir).activate()
  sfMain.getRange("M"+yeniSatir).setValue("ST")
  if(!sfMain.getRange("B"+yeniSatir).getValue())
    sfMain.getRange("B"+yeniSatir).setValue("FirmaAdı")

}

//Yazılan bilgilere barkod ve etiket çıkarmak için
function kargoEtiket(barSayac = null){ 
  //Etiket Aralığı: 275247000000X-275247999999X 1.000.000 (bir milyon) adet
  //Bizim aralığımız: 275247990001X - 275247999999X
  //Beylikdüzü depo için başlangıç: 275247980000
  //Merkez için başlangıç: 275247990000

  //Barkod sayacını kontrol edelim 
  if(!barSayac)
    throw("Sayaç hatası. Yönetici ile görüşünüz")
  

  //Gereken parametreler
  var zaman = Utilities.formatDate(new Date(), "GMT+3", "yyMMdd-HHmmss")
  barSayac = Number(barSayac)
  var sayac=0

  //Data aralığımızı alalım ve barkod çıkarılacak satırları filtreleyelim.(W,X,Y,Z sütunları kontrolü)
  let data = sfMain.getDataRange().getValues().filter((row,idx)=>{
    row.unshift(idx)
    return row[23] == "✅" && row[24] == "✅" && row[25] == "✅" && row[26] == ""
  })
  
  // Filtrelenen satırlarda gezelim
  data.forEach(row=>{
    
    //Barkodu hesaplayalım
    let tekler = 0, ciftler = 0, tmpEk = null
    
    let tmpBarkod = "275247" + (980000 + barSayac)
    tmpBarkod.split("").forEach(function(r,rid){
      rid % 2 == 0 ? tekler += Number(r) : ciftler += (3 * Number(r))
      //console.log({r,idx,tekler,ciftler})
    })
    tmpEk = (ciftler + tekler) % 10
    if( tmpEk > 0 ){ tmpEk = 10 - tmpEk}

    //Barkod değerini yazalım
    sfMain.getRange("L"+(row[0]+1)).setValue(tmpBarkod + tmpEk)
    sfMain.getRange("Z"+(row[0]+1)).setValue(zaman)

    let satirData = `100449199##201107091530##${row[0]+1}##${tmpBarkod+tmpEk}##${row[6]}##${row[14]}######${row[13]}######${row[8]}` +
      `##ÇAMLICA BASIM YAYIN VE TİCARET A.Ş.##BAĞLAR MAH.MİMARSİNAN CAD.NO:52-54 GÜNEŞLI BAĞCILAR/İSTANBUL##342##04##` + 
      `${row[2]}##${row[3]}##${row[10]}##${row[11]}##` +
      `####${row[22]}####` +
      `##${row[19]}##${row[19]}##${row[21]}##${row[7]}######??`

    sfMain.getRange("AA"+(row[0]+1)).setValue(satirData)
    //sfMain.getRange("AA"+(row[0]+1)).setBackgroundRGB(217,243,186)
    
    //Barkod sayacını arttıralım
    barSayac++
    sayac++
  })


  if(sayac == 0)
    throw("Çıkarılacak barkod yok!!")
    
  //Verdiğimiz zaman parametresini gönderelim
  sfMain.getRange("Z1").setValue(zaman)

  //Etiket sayfasından fazla satırları gizleyelim
  sfEtiket.showRows(1,sfEtiket.getMaxRows())
  sfEtiket.hideRows( 15*sayac+1,(sfEtiket.getMaxRows() - 15*sayac) )
  sfEtiket.activate()

  ////Mesaj verelim
  SpreadsheetApp.getActiveSpreadsheet().toast(`Çıkarılan Barkod:${sayac}`,"BİLGİ")

  //Kaldığımız barkodu dönelim
  return barSayac

}

//Log Yazmak için
function logYaz(param1, param2){  //Log yazalım
  sfLog.appendRow([new Date(),logYaz.caller.name, param1, param2])
}

//Etiket girişi için hazırlanalım
function etiketGirisi(satirNo = null){

  //İşlem yapılacak satır
  let yeniSatir = satirNo || sfMain.getLastRow()+1
  
  //Yeni Satıra konumlanıp, Varsayılan bilgileri yazalım
  sfMain.getRange("B"+yeniSatir).activate()
  sfMain.getRange("M"+yeniSatir).setValue("ST")
  if(!sfMain.getRange("B"+yeniSatir).getValue())
    sfMain.getRange("B"+yeniSatir).setValue("FirmaAdı")

}
