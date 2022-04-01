# LibreOffice-Türkçe-Dilinde-Makro-Kaynakları
LibreBasic İle LibreOffice Üzerinde Yapılacak Makro Kaynakları için Bir Rehberdir.

Kaynak: [Libre Office 7.2 Başlangıç Kılavuzu](https://documentation.libreoffice.org/assets/Uploads/Documentation/tr/GS7.2/LibreOffice-Balangc-Klavuzu.pdf)

 //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                         PROGRAM_1:                    TANITIM
 Sayfa numarası: 446 'da            ///////////////////////////////////////////////////////////////////////////////////        30/03/2022

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
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                        PROGRAM_2:                    MAKRO KAYDET
 Sayfa 452 'den devam et.            //////////////////////////////////////////////////////////////////////////////////////        31/03/2022

```
REM  *****  BASIC  *****


              'LibreOffice Basic makro dilinde tanımlamaları Türkçe alfabesinde yazarsanız
                'derleyici size hata mesajı gönderir. Bunun yerine Türkçe kelimeleri
                ' İngiliz alfabesinde yazmanız gerekiyor.
                 ' Bu uygulamayı tasarladıkları için ve şimdilik bunu nasıl değştireceğimizi
                ' bilmediğimiz için kurallara sadık kalalım.
              'Ancak kelimeler arasında boşluk bırakmak yerine alt çizgiyi "_" kullanmalısınız.


REM  *****  BASIC  *****


rem -------------------------------------------------------------

rem Hep hatırlayalım, kodlarımız her zaman main olarak yazılan yerin arasında çalışır
Sub Main

hazir_Veriyi_Yaz      rem hazir_Veriyi_Yaz ile yazılan yerde bizim yapılmasını istediklerimizi tanımladık.

End Sub


sub hazir_Veriyi_Yaz  REM Biz, gerekenleri yaptıktan sonra aşağıdakileri program kendisi yazıyor.
                     Rem Yani, kendimiz klavyede yazmadık.
rem ----------------------------------------------------------------------
rem define variables
dim belge   as object
dim dispatcher as object
rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

rem ----------------------------------------------------------------------
dim args1(0) as new com.sun.star.beans.PropertyValue
args1(0).Name = "Bold"
args1(0).Value = true

dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args1())

rem ----------------------------------------------------------------------
dim args2(0) as new com.sun.star.beans.PropertyValue
args2(0).Name = "Text"
args2(0).Value = "Bu uzun makro ."

dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args2())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:InsertPara", "", 0, Array())

rem ----------------------------------------------------------------------
dim args4(0) as new com.sun.star.beans.PropertyValue
args4(0).Name = "Text"
args4(0).Value = "."

dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args4())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:InsertPara", "", 0, Array())

rem ----------------------------------------------------------------------
dim args6(0) as new com.sun.star.beans.PropertyValue
args6(0).Name = "Text"
args6(0).Value = "."

dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args6())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:InsertPara", "", 0, Array())

rem ----------------------------------------------------------------------
dim args8(0) as new com.sun.star.beans.PropertyValue
args8(0).Name = "Text"
args8(0).Value = "."

dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args8())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:InsertPara", "", 0, Array())

rem ----------------------------------------------------------------------
dim args10(0) as new com.sun.star.beans.PropertyValue
args10(0).Name = "Text"
args10(0).Value = "."

dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args10())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:InsertPara", "", 0, Array())

rem ----------------------------------------------------------------------
dim args12(0) as new com.sun.star.beans.PropertyValue
args12(0).Name = "Text"
args12(0).Value = ".."

dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args12())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:InsertPara", "", 0, Array())

rem ----------------------------------------------------------------------
dim args14(0) as new com.sun.star.beans.PropertyValue
args14(0).Name = "Text"
args14(0).Value = "."

dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args14())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:InsertPara", "", 0, Array())

rem ----------------------------------------------------------------------
dim args16(0) as new com.sun.star.beans.PropertyValue
args16(0).Name = "Text"
args16(0).Value = "."

dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args16())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:InsertPara", "", 0, Array())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:InsertPara", "", 0, Array())

rem ----------------------------------------------------------------------
dim args19(0) as new com.sun.star.beans.PropertyValue
args19(0).Name = "Text"
args19(0).Value = "."

dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args19())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:InsertPara", "", 0, Array())

rem ----------------------------------------------------------------------
dim args21(0) as new com.sun.star.beans.PropertyValue
args21(0).Name = "Text"
args21(0).Value = "."

dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args21())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:InsertPara", "", 0, Array())

rem ----------------------------------------------------------------------
dim args23(0) as new com.sun.star.beans.PropertyValue
args23(0).Name = "Text"
args23(0).Value = "."

dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args23())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:InsertPara", "", 0, Array())

rem ----------------------------------------------------------------------
dim args25(0) as new com.sun.star.beans.PropertyValue
args25(0).Name = "Text"
args25(0).Value = ". "

dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args25())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:SwBackspace", "", 0, Array())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:SwBackspace", "", 0, Array())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:SwBackspace", "", 0, Array())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:SwBackspace", "", 0, Array())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:SwBackspace", "", 0, Array())

rem ----------------------------------------------------------------------
dim args31(0) as new com.sun.star.beans.PropertyValue
args31(0).Name = "Text"
args31(0).Value = "metinmakro kullanılarakkaydet özelliği  oluşturuldu DENR"

dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args31())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:ShiftBackspace", "", 0, Array())

rem ----------------------------------------------------------------------
dim args33(0) as new com.sun.star.beans.PropertyValue
args33(0).Name = "Text"
args33(0).Value = "EME sayısı :1 . "

dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args33())


end sub
```
 ////////////////////////////////////////////////////////////////////////////////////////////////////////
                                       PROGRAM_3:       VERİLERİ OTOMATİK OLARAK BİR TABLOYA KATMAK
 Sayfa 455 'ten devam et.            ////////////////////////////////////////////////////////////////////////////////////// Tarih: 
