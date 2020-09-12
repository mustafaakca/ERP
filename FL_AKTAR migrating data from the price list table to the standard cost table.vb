--This code is migrating data from the price list table to the standard cost table to help the user calculate the expected price correctly

Sub Makro1()
 if trim(Doc.GetGecoValue("STMTYP","KOD","TK",d7c_ErrMsgNo)) = "" then
  ret = GrupKoduYarat("STMTYP","TK","Teklif Endeksi")
 end if
 if ret <> 0 then
  k = Doc.d7MsgBox("TK Maliyet endeksi yaratılamadı. İşlem kesilmiştir.","dinamo",MB_OK)
  exit sub
 end if
 if trim(Doc.GetGecoValue("STMMUS","KOD","HM",d7c_ErrMsgNo)) = "" then
  ret = GrupKoduYarat("STMMUS","HM","Hammadde Maliyeti")
 end if
 if ret <> 0 then
  k = Doc.d7MsgBox("HM masraf unsuru yaratılamadı. İşlem kesilmiştir.","dinamo",MB_OK)
  exit sub
 end if

 Set Doc1 = Doc.CreateDocumentNoUI("STDMH0","")
 Set STDMH0E = Doc1.GetTableObject("STDMH0E")
 Set STDMH0T = Doc1.GetTableObject("STDMH0T")
 For i = 1 to STOK48T.GetRecCount ()
  if trim(STOK48T.KOD(i)) <> "" then
   StokKodu = STOK48T.KOD(i)
   rows = Doc.SelectEQ("STDMH0E","KOD",StokKodu)
   if rows = 0 then
    ret = Doc1.New_Voucher()    
    STDMH0E.STDMALTYPE = "TK"
    STDMH0E.KOD = StokKodu
    STDMH0T.AddRow
    satirNo = STDMH0T.GetRecCount
   else
    set ST = Doc.GetTableObject("AU_STDMH0E_SELECT")
    For j = 1 to ST.GetRecCount()
     if trim(ST.STDMALTYPE(j)) = "TK" then
      EvrakNo = ST.EVRAKNO(j)
     end if
    Next
    if trim(EvrakNo) <> "" then
     ret = Doc1.Load_Voucher(EvrakNo)
     if ret <> 0 then
      k = Doc.d7MsgBox(trim(StokKodu) & ": Standart Maliyet tablosu yüklenemedi. İşlem kesilmiştir.","dinamo",MB_OK)
      exit sub
     end if
     satirNo = satirBul(STDMH0T)
     if satirNo = 0 then
      STDMH0T.AddRow
      satirNo = STDMH0T.GetRecCount
     End if
    else
     ret = Doc1.New_Voucher()    
     STDMH0E.STDMALTYPE = "TK"
     STDMH0E.KOD = StokKodu
     STDMH0T.AddRow
     satirNo = STDMH0T.GetRecCount
    end if
   end if
   STDMH0T.VALIDAFTERTARIH(satirNo) = Doc.StrToNorDate(Doc.BuGun)
   STDMH0T.UNSUR(satirNo) = "HM"
   STDMH0T.TUTAR (satirNo) = STOK48T.PRICE(i)
   STDMH0T.PARABIRIMI(satirNo) = STOK48T.PRICEUNIT(i)
   ret = Doc1.Save_Voucher()
   if ret <> 0 then
    k = Doc.d7MsgBox(trim(StokKodu) & ": Standart Maliyet tablosu kaydedilemedi. İşlem kesilmiştir.","dinamo",MB_OK)
    ret = Doc1.LoadEmpty_Voucher()
    exit sub
   end if
   ret = Doc1.LoadEmpty_Voucher()
  end if
 Next
End Sub
Function GrupKoduYarat(GrupKoduTipi,GrupKodu,GrupAciklamasi)
 Set Doc1 = Doc.CreateDocumentNoUI ("GECO10","")
 Set GT = Doc1.GetTableObject("GECOUST")
 ret = Doc1.Load_Voucher(GrupKoduTipi)
 if ret <> 0 then
  GrupKoduYarat = 1
  ret = Doc1.LoadEmpty_Voucher()
  exit Function
 end if
 GT.AddRow
 satirNo = GT.GetRecCount
 GT.KOD(satirNo) = GrupKodu
 GT.AD(satirNo) = GrupAciklamasi
 ret = Doc1.Save_Voucher()
 if ret <> 0 then
  GrupKoduYarat = 1
 else
  GrupKoduYarat = 0
 end if
 ret = Doc1.LoadEmpty_Voucher()
End Function
Function satirBul(STDMH0T)
 satirBul = 0
 For i = 1 to STDMH0T.GetRowCount()
  if STDMH0T.VALIDAFTERTARIH(i) = Doc.BuGun then
   satirBul = i
   exit function
  end if
 Next
End Function