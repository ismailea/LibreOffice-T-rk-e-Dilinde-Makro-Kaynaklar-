# LibreOffice-Türkçe-Dilinde-Makro-Kaynakları
LibreBasic İle LibreOffice Üzerinde Yapılacak Makro Kaynakları için Bir Rehberdir.

Kaynak: [Libre Office 7.2 Başlangıç Kılavuzu](https://documentation.libreoffice.org/assets/Uploads/Documentation/tr/GS7.2/LibreOffice-Balangc-Klavuzu.pdf)

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Sayfa numarası: 446 'da            //////////////////////////////////////////////////////////////////////////////////////        30/03/2022

```
REM  *****  BASIC  *****
REM Bu ifade tek satırlık bir yorum ya da açıklama yapmamıza olanak tanır.
rem bu ifade de teksatırlık bir yorum belirtir.


Sub merhabaMakro
Print "Merhaba"
End Sub


Sub Main

merhabaMakro 'Bu bir yorum parantezidir. Program kesme işaretinden sonrakileri derlemez.
              ' Açıklamalar için kullanılır. Sadece bir satır için geçerlidir.

End Sub
```
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Sayfa 452 'den devam et.            //////////////////////////////////////////////////////////////////////////////////////        31/03/2022

```
REM  *****  BASIC  *****


sub benim_Adimi_Gir   'LibreOffice Basic makro dilinde tanımlamaları Türkçe alfabesinde yazarsanız
                'derleyici size hata mesajı gönderir. Bunun yerine Türkçe kelimeleri
                ' İngiliz alfabesinde yazmanız gerekiyor.
                 ' Bu uygulamayı tasarladıkları için ve şimdilik bunu nasıl değştireceğimizi
                 ' bilmediğimiz için kurallara sadık kalalım.
                 'Ancak kelimeler arasında boşluk bırakmak yerine alt çizgiyi "_" kullanmalısınız.

rem -------------------------------------------------------------
rem değişkenleri tanımla
dim belge as object
dim sevk_memuru as object
rem -------------------------------------------------------------

rem belgeye erişim elde et
belge = ThisComponent.CurrentController.Frame
sevkMemuru = createUnoService("com.sun.star.frame.DispatchHelper")
rem -------------------------------------------------------------

dim kalemler(0) as new com.sun.star.beans.PropertyValue
kalemler(0).Name = "Sizden_bir_kac_bilgi_alacagiz."
kalemler(0).Value = "Adinizi_yaziniz: "
sevkMemuru.executeDispatch(belge, ".uno:InsertText", "", 0, kalemler())

end sub
rem -------------------------------------------------------------

rem Hep hatırlayalım, kodlarımız her zaman main olarak yazılan yerin arasında çalışır
Sub Main

benim_Adimi_Gir      'benim_Adimi_Gir ile yazılan yerde bizim yapılmasını istediklerimizi tanımladık.

End Sub
```
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
