Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration

Public Class Auto_Reoder

    Dim conRE As String = ConfigurationManager.ConnectionStrings("DBRE_ConnectionString").ConnectionString
    Dim conRT As String = ConfigurationManager.ConnectionStrings("DBRT_ConnectionString").ConnectionString
    Dim conPY As String = ConfigurationManager.ConnectionStrings("DBPY_ConnectionString").ConnectionString
    Dim conPK As String = ConfigurationManager.ConnectionStrings("DBPK_ConnectionString").ConnectionString
    Dim conSW As String = ConfigurationManager.ConnectionStrings("DBSW_ConnectionString").ConnectionString
    Dim conNR As String = ConfigurationManager.ConnectionStrings("DBNR_ConnectionString").ConnectionString
    Dim conHR As String = ConfigurationManager.ConnectionStrings("DBHR_ConnectionString").ConnectionString

    Dim ConHs As String = String.Empty

    Dim success As Boolean

    Dim at As Integer = 9010

    ' Dim at As Integer = 305


    Private Function ConSrv(ByVal srv As String) As String

        ConHs = ""
        If srv.Contains("HSRE") Then
            ConHs = conRE
        ElseIf srv.Contains("HSRT") Then
            ConHs = conRT
        ElseIf srv.Contains("HSPK") Then
            ConHs = conPK
        ElseIf srv.Contains("HSPY") Then
            ConHs = conPY
        ElseIf srv.Contains("HSSW") Then
            ConHs = conSW
        ElseIf srv.Contains("HSNR") Then
            ConHs = conNR
        End If
        Return ConHs
    End Function

    Private Sub del()

        Dim sqlDel As String
        '  sqlDel = "delete [HS101].[dbo].[Ord_AutoPrd]"
        sqlDel = "delete [HS101].[dbo].[Ord_autoReoder] where srv in('HSSW','HSNR','HSRE')"
        Using conn As New SqlConnection(ConSrv("HSRE"))
            Try
                conn.Open()
                Dim cmd As SqlCommand = New SqlCommand(sqlDel, conn)
                cmd.CommandTimeout = 2220
                cmd.ExecuteNonQuery()
                conn.Close()
            Catch ex As Exception

            End Try

        End Using

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        at = at - 1
        lbxx.Text = at.ToString

        If (at = 9000) Then
            innitdb("HSSW")
        ElseIf (at = 7000) Then
            innitdb("HSNR")
        ElseIf (at = 6000) Then
            innitdb("HSRE")

            'ElseIf at = 0 Then
            '  Application.Exit()
        End If

        'If (at = 300) Then
        '    innitdb("HSRE")

        '    lbcc.Text = "success :" + success.ToString
        'ElseIf (at = 200) Then
        '    innitdb("HSRT")

        '    lbcc.Text = "success :" + success.ToString
        'ElseIf (at = 100) Then
        '    innitdb("HSPK")
        '    lbcc.Text = "success :" + success.ToString
        'End If



    End Sub

    Private Sub Auto_Reoder_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        del()
        Timer1.Start()

    End Sub




    Private Sub innitdb(ByVal srv As String)


        Dim sqlQryre As String = ""
        Dim lsrv As String = srv.Substring(2, 2)

        'sqlQryre = "INSERT INTO  [HS101].[dbo].[test_readexcel]  ([customerid] ,[city] ,[country] ,[postcode]) VALUES   ('a1','roied','thai','45000')"
        '  sqlQryre = " INSERT INTO [HS101].[dbo].[Ord_autoReoder](apcode ,apname ,skucode,skuname ,utqname ,salesum  ,salemax  ,saleday ,sqty,rqty ,srv,udate,dstk,iccode) SELECT  ap_code,ap_name ,sk.SKU_CODE,sk.SKU_NAME,utq_name,(SELECT ISNULL(  sum( ABS (skm_qty)),0.00) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE  dt.DT_PROPERTIES in('302','307') AND sku_code = sk.SKU_CODE  AND DATEDIFF(day,di.DI_DATE,getdate()) between 0 and 365  ) salesum , (SELECT ISNULL(  max( ABS (skm_qty)),0.00)FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE  dt.DT_PROPERTIES in('302','307') AND sku_code =  sk.SKU_CODE  AND DATEDIFF(day,di.DI_DATE,getdate()) between 0 and 365  ) AS salemax,( SELECT ISNULL( max(  DATEDIFF(day,di.DI_DATE,getdate())),0) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE  dt.DT_PROPERTIES in('302','307') AND sku_code = sk.SKU_CODE  AND DATEDIFF(day,di.DI_DATE,getdate()) between 0 and 365  ) AS saleday,ISNULL( SUM(SKM_QTY),0) SQTY, ISNULL( SH_QTY,0) RQTY,'" + srv + "' AS srv,CONVERT(smalldatetime, GETDATE()),(	SELECT ISNULL(  sum( ABS (skm_qty)),0.00) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE  dt.DT_PROPERTIES in('302','307') AND sku_code = sk.SKU_CODE  AND DATEDIFF(day,di.DI_DATE,getdate()) between 0 and 180  ) AS dstk,ICCAT_CODE FROM APFILE INNER JOIN SKUAP ON APFILE.AP_KEY = SKP_AP  INNER JOIN dbo.SKUMASTER sk ON  SKP_SKU = sk.SKU_KEY LEFT JOIN HS_BP_BAL" + lsrv + " ON TRD_SH_CODE = sk.SKU_CODE     INNER JOIN SKUMOVE ON skm_sku=sk.sku_key    INNER JOIN dbo.ICCAT ON SKU_ICCAT = ICCAT_KEY INNER JOIN dbo.UOFQTY ON SKU_S_UTQ = UTQ_KEY INNER JOIN dbo.BRAND ON SKU_BRN = BRN_KEY LEFT JOIN SKUALT ON  SKU_SKUALT= SKUALT_KEY LEFT JOIN  dbo.ICCOLOR ON SKU_ICCOLOR  = ICCOLOR_KEY LEFT JOIN dbo.ICSIZE  ON SKU_ICSIZE= ICSIZE_KEY LEFT JOIN dbo.ICGL ON SKU_ICGL= ICGL_KEY LEFT JOIN dbo.ICPRT ON SKU_ICPRT = ICPRT_KEY LEFT JOIN dbo.ICDEPT ON SKU_ICDEPT = ICDEPT_KEY    WHERE SKUAP.SKP_DEFAULT = 'Y' and SKU_ENABLE = 'Y' and  left(sku_code,1)  not in('a','9','Z','N')  and LEFT(sku_code,2)not in('01','02','59','50','40','30','41')  and  right(sku_code,2) <> 'NO' and right(AP_CODE,2) <> 'NO'   and right(AP_NAME,2) <> 'NO'   GROUP BY  ap_code,ap_name ,sku_code,sku_name,utq_name ,SH_QTY,ICCAT_CODE "

        '5/6/18
        '  sqlQryre = "INSERT INTO [HS101].[dbo].[Ord_autoReoder](apcode ,apname ,skucode,skuname ,utqname ,salesum  ,salemax  ,saleday ,sqty,rqty ,srv,udate,dstk,iccode) SELECT   ISNULL( (SELECT top 1  AP_CODE FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY  LEFT JOIN  APDETAIL AL ON di.DI_KEY = AL.APD_DI  INNER JOIN APFILE AF ON AL.APD_AP = AF.AP_KEY WHERE   sku_code = sk.SKU_CODE   and  LEFT(di.DI_REF,3) ='IBP' and di.di_date = (  SELECT MAX(DI_DATE) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE   sku_code = sk.SKU_CODE    and  LEFT(di.DI_REF,3) ='IBP') ),AP_CODE)   ap_code, ISNULL( (SELECT top 1  AP_NAME FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY  LEFT JOIN  APDETAIL AL ON di.DI_KEY = AL.APD_DI  INNER JOIN APFILE AF ON AL.APD_AP = AF.AP_KEY WHERE   sku_code = sk.SKU_CODE   and  LEFT(di.DI_REF,3) ='IBP' and di.di_date = (  SELECT MAX(DI_DATE) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE   sku_code = sk.SKU_CODE    and  LEFT(di.DI_REF,3) ='IBP') ),AP_NAME)  ap_name ,sk.SKU_CODE,sk.SKU_NAME,utq_name,(SELECT ISNULL(  sum( ABS (skm_qty)),0.00) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE  dt.DT_PROPERTIES in('302','307') AND sku_code = sk.SKU_CODE  AND DATEDIFF(day,di.DI_DATE,getdate()) between 0 and 365  ) salesum ,(SELECT ISNULL(  max( ABS (skm_qty)),0.00)FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE  dt.DT_PROPERTIES in('302','307') AND sku_code =  sk.SKU_CODE  AND DATEDIFF(day,di.DI_DATE,getdate()) between 0 and 365  ) AS salemax,ISNULL( ( SELECT  CASE WHEN  datediff(day,  (SELECT  min(di.DI_DATE) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY    WHERE   sku_code = sk.SKU_CODE   and  LEFT(di.DI_REF,3) ='IBP')  ,getdate()) >= 365 THEN 365 ELSE  datediff(day,  (SELECT  min(di.DI_DATE) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY    WHERE   sku_code = sk.SKU_CODE   and LEFT(di.DI_REF,3) ='IBP')  ,getdate()) END) , 0 )as saleday, ISNULL( SUM(SKM_QTY),0) SQTY, ISNULL( SH_QTY,0) RQTY,'" + srv + "' AS srv,CONVERT(smalldatetime, GETDATE()),(	SELECT ISNULL(  sum( ABS (skm_qty)),0.00) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE  dt.DT_PROPERTIES in('302','307') AND sku_code = sk.SKU_CODE  AND DATEDIFF(day,di.DI_DATE,getdate()) between 0 and 180  ) AS dstk,ICCAT_CODE  FROM APFILE INNER JOIN SKUAP ON APFILE.AP_KEY = SKP_AP  INNER JOIN dbo.SKUMASTER sk ON  SKP_SKU = sk.SKU_KEY LEFT JOIN HS_BP_BAL" + lsrv + " ON TRD_SH_CODE = sk.SKU_CODE     INNER JOIN SKUMOVE ON skm_sku=sk.sku_key    INNER JOIN dbo.ICCAT ON SKU_ICCAT = ICCAT_KEY INNER JOIN dbo.UOFQTY ON SKU_S_UTQ = UTQ_KEY INNER JOIN dbo.BRAND ON SKU_BRN = BRN_KEY LEFT JOIN SKUALT ON  SKU_SKUALT= SKUALT_KEY LEFT JOIN  dbo.ICCOLOR ON SKU_ICCOLOR  = ICCOLOR_KEY LEFT JOIN dbo.ICSIZE  ON SKU_ICSIZE= ICSIZE_KEY LEFT JOIN dbo.ICGL ON SKU_ICGL= ICGL_KEY LEFT JOIN dbo.ICPRT ON SKU_ICPRT = ICPRT_KEY LEFT JOIN dbo.ICDEPT ON SKU_ICDEPT = ICDEPT_KEY    WHERE SKUAP.SKP_DEFAULT = 'Y' and SKU_ENABLE = 'Y' and  left(sku_code,1)  not in('a','9','Z','N')  and LEFT(sku_code,2)not in('01','02','59','50','40','30','41')  and  right(sku_code,2) <> 'NO' and right(AP_CODE,2) <> 'NO'   and right(AP_NAME,2) <> 'NO'  GROUP BY  ap_code,ap_name ,sku_code,sku_name,utq_name ,SH_QTY,ICCAT_CODE "

        'ชื้อล่าสุด ใช้
        ' sqlQryre = "INSERT INTO [HS101].[dbo].[Ord_autoReoder](apcode ,apname ,skucode,skuname ,utqname ,salesum  ,salemax  ,saleday ,sqty,rqty ,srv,udate,dstk,iccode) SELECT   ISNULL( (SELECT top 1  AP_CODE FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY  LEFT JOIN  APDETAIL AL ON di.DI_KEY = AL.APD_DI  INNER JOIN APFILE AF ON AL.APD_AP = AF.AP_KEY WHERE   sku_code = sk.SKU_CODE   and  LEFT(di.DI_REF,3) ='IBP' and di.di_date = (  SELECT MAX(DI_DATE) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE   sku_code = sk.SKU_CODE    and  LEFT(di.DI_REF,3) ='IBP') ),AP_CODE)   ap_code, ISNULL( (SELECT top 1  AP_NAME FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY  LEFT JOIN  APDETAIL AL ON di.DI_KEY = AL.APD_DI  INNER JOIN APFILE AF ON AL.APD_AP = AF.AP_KEY WHERE   sku_code = sk.SKU_CODE   and  LEFT(di.DI_REF,3) ='IBP' and di.di_date = (  SELECT MAX(DI_DATE) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE   sku_code = sk.SKU_CODE    and  LEFT(di.DI_REF,3) ='IBP') ),AP_NAME)  ap_name,sk.SKU_CODE,sk.SKU_NAME,utq_name,(SELECT ISNULL(  sum( ABS (skm_qty)),0.00) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE  dt.DT_PROPERTIES in('302','307') AND sku_code = sk.SKU_CODE  AND DATEDIFF(day,di.DI_DATE,getdate()) >= 365  ) salesum ,(SELECT ISNULL(  max( ABS (skm_qty)),0.00)FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE dt.DT_PROPERTIES in('302','307') AND sku_code =  sk.SKU_CODE  AND DATEDIFF(day,di.DI_DATE,getdate()) >= 365 ) AS salemax,ISNULL( ( SELECT  CASE WHEN  datediff(day,  (SELECT  min(di.DI_DATE) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY    WHERE   sku_code = sk.SKU_CODE   and  LEFT(di.DI_REF,3) ='IBP')  ,getdate()) >= 365 THEN 365 ELSE  datediff(day,  (SELECT  min(di.DI_DATE) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY    WHERE   sku_code = sk.SKU_CODE   and LEFT(di.DI_REF,3) ='IBP')  ,getdate()) END) , 0 )as saleday, ISNULL( SUM(SKM_QTY),0) SQTY, ISNULL( SH_QTY,0) RQTY,'" + srv + "' AS srv,CONVERT(smalldatetime, GETDATE()),(	SELECT ISNULL(  sum( ABS (skm_qty)),0.00) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE  dt.DT_PROPERTIES in('302','307') AND sku_code = sk.SKU_CODE  AND DATEDIFF(day,di.DI_DATE,getdate()) between 0 and 180  ) AS dstk,ICCAT_CODE  FROM APFILE INNER JOIN SKUAP ON APFILE.AP_KEY = SKP_AP  INNER JOIN dbo.SKUMASTER sk ON  SKP_SKU = sk.SKU_KEY LEFT JOIN HS_BP_BAL" + lsrv + " ON TRD_SH_CODE = sk.SKU_CODE     INNER JOIN SKUMOVE ON skm_sku=sk.sku_key    INNER JOIN dbo.ICCAT ON SKU_ICCAT = ICCAT_KEY INNER JOIN dbo.UOFQTY ON SKU_S_UTQ = UTQ_KEY INNER JOIN dbo.BRAND ON SKU_BRN = BRN_KEY LEFT JOIN SKUALT ON  SKU_SKUALT= SKUALT_KEY LEFT JOIN  dbo.ICCOLOR ON SKU_ICCOLOR  = ICCOLOR_KEY LEFT JOIN dbo.ICSIZE  ON SKU_ICSIZE= ICSIZE_KEY LEFT JOIN dbo.ICGL ON SKU_ICGL= ICGL_KEY LEFT JOIN dbo.ICPRT ON SKU_ICPRT = ICPRT_KEY LEFT JOIN dbo.ICDEPT ON SKU_ICDEPT = ICDEPT_KEY    WHERE SKUAP.SKP_DEFAULT = 'Y' and SKU_ENABLE = 'Y' and  left(sku_code,1)  not in('a','9','Z','N')  and LEFT(sku_code,2)not in('01','02','59','50','40','30','41')  and  right(sku_code,2) <> 'NO' and right(AP_CODE,2) <> 'NO'   and right(AP_NAME,2) <> 'NO'  GROUP BY  ap_code,ap_name ,sku_code,sku_name,utq_name ,SH_QTY,ICCAT_CODE "

        'ผู้จำหน่ายหลัก ไม่ใช้ชื้อล่าสุด

        ' sqlQryre = "INSERT INTO [HS101].[dbo].[Ord_autoReoder](apcode ,apname ,skucode,skuname ,utqname ,salesum  ,salemax  ,saleday ,sqty,rqty ,srv,udate,dstk,iccode) SELECT    ap_code,ap_name,sk.SKU_CODE,sk.SKU_NAME,utq_name,(SELECT ISNULL(  sum( ABS (skm_qty)),0.00) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE  dt.DT_PROPERTIES in('302','307') AND sku_code = sk.SKU_CODE  AND DATEDIFF(day,di.DI_DATE,getdate())  between 0 and 365   ) salesum ,(SELECT ISNULL(  max( ABS (skm_qty)),0.00)FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE dt.DT_PROPERTIES in('302','307') AND sku_code =  sk.SKU_CODE  AND DATEDIFF(day,di.DI_DATE,getdate())  between 0 and 365  ) AS salemax,ISNULL( ( SELECT  CASE WHEN  datediff(day,  (SELECT  min(di.DI_DATE) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY    WHERE   sku_code = sk.SKU_CODE   and  LEFT(di.DI_REF,3) ='IBP')  ,getdate()) >= 365 THEN 365 ELSE  datediff(day,  (SELECT  min(di.DI_DATE) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY    WHERE   sku_code = sk.SKU_CODE   and LEFT(di.DI_REF,3) ='IBP')  ,getdate()) END) , 0 )as saleday, ISNULL( SUM(SKM_QTY),0) SQTY, ISNULL( SH_QTY,0) RQTY,'" + srv + "' AS srv,CONVERT(smalldatetime, GETDATE()),(	SELECT ISNULL(  sum( ABS (skm_qty)),0.00) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE  dt.DT_PROPERTIES in('302','307') AND sku_code = sk.SKU_CODE  AND DATEDIFF(day,di.DI_DATE,getdate()) between 0 and 180  ) AS dstk,ICCAT_CODE  FROM APFILE INNER JOIN SKUAP ON APFILE.AP_KEY = SKP_AP  INNER JOIN dbo.SKUMASTER sk ON  SKP_SKU = sk.SKU_KEY LEFT JOIN HS_BP_BAL" + lsrv + " ON TRD_SH_CODE = sk.SKU_CODE     INNER JOIN SKUMOVE ON skm_sku=sk.sku_key    INNER JOIN dbo.ICCAT ON SKU_ICCAT = ICCAT_KEY INNER JOIN dbo.UOFQTY ON SKU_S_UTQ = UTQ_KEY INNER JOIN dbo.BRAND ON SKU_BRN = BRN_KEY LEFT JOIN SKUALT ON  SKU_SKUALT= SKUALT_KEY LEFT JOIN  dbo.ICCOLOR ON SKU_ICCOLOR  = ICCOLOR_KEY LEFT JOIN dbo.ICSIZE  ON SKU_ICSIZE= ICSIZE_KEY LEFT JOIN dbo.ICGL ON SKU_ICGL= ICGL_KEY LEFT JOIN dbo.ICPRT ON SKU_ICPRT = ICPRT_KEY LEFT JOIN dbo.ICDEPT ON SKU_ICDEPT = ICDEPT_KEY    WHERE SKUAP.SKP_DEFAULT = 'Y' and SKU_ENABLE = 'Y' and  left(sku_code,1)  not in('a','9','Z','N')  and LEFT(sku_code,2)not in('01','02','59','50','40','30','41')  and  right(sku_code,2) <> 'NO' and right(AP_CODE,2) <> 'NO'   and right(AP_NAME,2) <> 'NO'  GROUP BY  ap_code,ap_name ,sku_code,sku_name,utq_name ,SH_QTY,ICCAT_CODE "


        sqlQryre = " INSERT INTO [HS101].[dbo].[Ord_autoReoder](apcode ,apname ,skucode,skuname ,utqname ,salesum  ,salemax  ,saleday ,sqty,rqty ,srv,udate,dstk,iccode) SELECT    ISNULL( (SELECT top 1  AP_CODE FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY  LEFT JOIN  APDETAIL AL ON di.DI_KEY = AL.APD_DI  INNER JOIN APFILE AF ON AL.APD_AP = AF.AP_KEY WHERE   sku_code = sk.SKU_CODE   and  LEFT(di.DI_REF,3) ='IBP' and di.di_date = (  SELECT MAX(DI_DATE) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE   sku_code = sk.SKU_CODE    and  LEFT(di.DI_REF,3) ='IBP') ),AP_CODE)   ap_code, ISNULL( (SELECT top 1  AP_NAME FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY  LEFT JOIN  APDETAIL AL ON di.DI_KEY = AL.APD_DI  INNER JOIN APFILE AF ON AL.APD_AP = AF.AP_KEY WHERE   sku_code = sk.SKU_CODE   and  LEFT(di.DI_REF,3) ='IBP' and di.di_date = (  SELECT MAX(DI_DATE) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE   sku_code = sk.SKU_CODE    and  LEFT(di.DI_REF,3) ='IBP') ),AP_NAME)  ap_name,sk.SKU_CODE,sk.SKU_NAME,utq_name,(SELECT ISNULL(  sum( ABS (skm_qty)),0.00) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE  dt.DT_PROPERTIES in('302','307') AND sku_code = sk.SKU_CODE  AND DATEDIFF(day,di.DI_DATE,getdate())  between 0 and 365   ) salesum ,(SELECT ISNULL(  max( ABS (skm_qty)),0.00)FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE dt.DT_PROPERTIES in('302','307') AND sku_code =  sk.SKU_CODE  AND DATEDIFF(day,di.DI_DATE,getdate())  between 0 and 365  ) AS salemax, ISNULL( (SELECT max( DATEDIFF(day,di.DI_DATE,getdate()))  FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY  WHERE dt.DT_PROPERTIES in('302','307') AND sku_code = sk.SKU_CODE  AND DATEDIFF(day,di.DI_DATE,getdate()) between 0 and 365),0) saleday, ISNULL( SUM(SKM_QTY),0) SQTY, ISNULL( SH_QTY,0) RQTY,'" + srv + "' AS srv,CONVERT(smalldatetime, GETDATE()),(	SELECT ISNULL(  sum( ABS (skm_qty)),0.00) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE  dt.DT_PROPERTIES in('302','307') AND sku_code = sk.SKU_CODE  AND DATEDIFF(day,di.DI_DATE,getdate()) between 0 and 180  ) AS dstk,ICCAT_CODE  FROM APFILE INNER JOIN SKUAP ON APFILE.AP_KEY = SKP_AP  INNER JOIN dbo.SKUMASTER sk ON  SKP_SKU = sk.SKU_KEY LEFT JOIN HS_BP_BAL" + lsrv + " ON TRD_SH_CODE = sk.SKU_CODE     INNER JOIN SKUMOVE ON skm_sku=sk.sku_key    INNER JOIN dbo.ICCAT ON SKU_ICCAT = ICCAT_KEY INNER JOIN dbo.UOFQTY ON SKU_S_UTQ = UTQ_KEY INNER JOIN dbo.BRAND ON SKU_BRN = BRN_KEY LEFT JOIN SKUALT ON  SKU_SKUALT= SKUALT_KEY LEFT JOIN  dbo.ICCOLOR ON SKU_ICCOLOR  = ICCOLOR_KEY LEFT JOIN dbo.ICSIZE  ON SKU_ICSIZE= ICSIZE_KEY LEFT JOIN dbo.ICGL ON SKU_ICGL= ICGL_KEY LEFT JOIN dbo.ICPRT ON SKU_ICPRT = ICPRT_KEY LEFT JOIN dbo.ICDEPT ON SKU_ICDEPT = ICDEPT_KEY    WHERE SKUAP.SKP_DEFAULT = 'Y' and SKU_ENABLE = 'Y' and  left(sku_code,1)  not in('a','9','Z','N','0') and LEFT( ICCAT_CODE,1) not in('ถ','ร','ค')  GROUP BY  ap_code,ap_name ,sku_code,sku_name,utq_name ,SH_QTY,ICCAT_CODE"

        '27/9/18  last


        'sqlQryre = " SELECT    ISNULL( (SELECT top 1  AP_CODE FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY  LEFT JOIN  APDETAIL AL ON di.DI_KEY = AL.APD_DI  INNER JOIN APFILE AF ON AL.APD_AP = AF.AP_KEY WHERE   sku_code = sk.SKU_CODE   and  LEFT(di.DI_REF,3) ='IBP' and di.di_date = (  SELECT MAX(DI_DATE) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE   sku_code = sk.SKU_CODE    and  LEFT(di.DI_REF,3) ='IBP') ),AP_CODE)   ap_code, ISNULL( (SELECT top 1  AP_NAME FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY  LEFT JOIN  APDETAIL AL ON di.DI_KEY = AL.APD_DI  INNER JOIN APFILE AF ON AL.APD_AP = AF.AP_KEY WHERE   sku_code = sk.SKU_CODE   and  LEFT(di.DI_REF,3) ='IBP' and di.di_date = (  SELECT MAX(DI_DATE) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE   sku_code = sk.SKU_CODE    and  LEFT(di.DI_REF,3) ='IBP') ),AP_NAME)  ap_name,sk.SKU_CODE,sk.SKU_NAME,utq_name,(SELECT ISNULL(  sum( ABS (skm_qty)),0.00) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE  dt.DT_PROPERTIES in('302','307') AND sku_code = sk.SKU_CODE  AND DATEDIFF(day,di.DI_DATE,getdate())  between 0 and 365   ) salesum ,(SELECT ISNULL(  max( ABS (skm_qty)),0.00)FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE dt.DT_PROPERTIES in('302','307') AND sku_code =  sk.SKU_CODE  AND DATEDIFF(day,di.DI_DATE,getdate())  between 0 and 365  ) AS salemax, ISNULL( (SELECT max( DATEDIFF(day,di.DI_DATE,getdate()))  FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY  WHERE dt.DT_PROPERTIES in('302','307') AND sku_code = sk.SKU_CODE  AND DATEDIFF(day,di.DI_DATE,getdate()) between 0 and 365),0) saleday, ISNULL( SUM(SKM_QTY),0) SQTY, ISNULL( SH_QTY,0) RQTY,'" + srv + "' AS srv,CONVERT(smalldatetime, GETDATE()),(	SELECT ISNULL(  sum( ABS (skm_qty)),0.00) FROM  SKUMOVE sm inner join SKUMASTER  ON  skm_sku = SKU_KEY  INNER JOIN DOCINFO di ON sm.SKM_DI = di.DI_KEY INNER JOIN DOCTYPE dt ON di.DI_DT=dt.DT_KEY   WHERE  dt.DT_PROPERTIES in('302','307') AND sku_code = sk.SKU_CODE  AND DATEDIFF(day,di.DI_DATE,getdate()) between 0 and 180  ) AS dstk,ICCAT_CODE  FROM APFILE INNER JOIN SKUAP ON APFILE.AP_KEY = SKP_AP  INNER JOIN dbo.SKUMASTER sk ON  SKP_SKU = sk.SKU_KEY LEFT JOIN HS_BP_BAL" + lsrv + " ON TRD_SH_CODE = sk.SKU_CODE     INNER JOIN SKUMOVE ON skm_sku=sk.sku_key    INNER JOIN dbo.ICCAT ON SKU_ICCAT = ICCAT_KEY INNER JOIN dbo.UOFQTY ON SKU_S_UTQ = UTQ_KEY INNER JOIN dbo.BRAND ON SKU_BRN = BRN_KEY LEFT JOIN SKUALT ON  SKU_SKUALT= SKUALT_KEY LEFT JOIN  dbo.ICCOLOR ON SKU_ICCOLOR  = ICCOLOR_KEY LEFT JOIN dbo.ICSIZE  ON SKU_ICSIZE= ICSIZE_KEY LEFT JOIN dbo.ICGL ON SKU_ICGL= ICGL_KEY LEFT JOIN dbo.ICPRT ON SKU_ICPRT = ICPRT_KEY LEFT JOIN dbo.ICDEPT ON SKU_ICDEPT = ICDEPT_KEY    WHERE SKUAP.SKP_DEFAULT = 'Y' and SKU_ENABLE = 'Y' and  left(sku_code,1)  not in('a','9','Z','N','0') and LEFT( ICCAT_CODE,1) not in('ถ','ร','ค')  GROUP BY  ap_code,ap_name ,sku_code,sku_name,utq_name ,SH_QTY,ICCAT_CODE"

        txtMsg.Text = sqlQryre


        Using conn As New SqlConnection(ConSrv(srv))


            ' lbSrv.Text = sqlQryre
            Dim cmd As SqlCommand = New SqlCommand(sqlQryre, conn)
            conn.Open()
            Try
                cmd.CommandTimeout = 18500
                cmd.ExecuteNonQuery()
                lbSrv.Text += srv + "_" + lbxx.Text + "_"
                txtMsg.Text = sqlQryre


            Catch ex As Exception
                lbSrv.Text += ex.ToString
            Finally
                conn.Close()
            End Try


            If cmd.ExecuteNonQuery() > 0 Then
                success = True
            Else
                success = False
            End If


        End Using




    End Sub


End Class
