# LibreOffice-Türkçe-Dilinde-Makro-Kaynakları
LibreBasic İle LibreOffice Üzerinde Yapılacak Makro Kaynakları için Bir Rehberdir.

Kaynak: [Libre Office 7.2 Başlangıç Kılavuzu](https://documentation.libreoffice.org/assets/Uploads/Documentation/tr/GS7.2/LibreOffice-Balangc-Klavuzu.pdf)
Sayfa numarası: 446 'da

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
