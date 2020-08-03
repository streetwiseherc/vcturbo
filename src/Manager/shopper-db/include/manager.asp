<% REM -- This file gets included by all Store Manager pages

    Set mscsPage            = Server.CreateObject("Commerce.Page")

    REM -- ADO command types
    adCmdText       = 1 
    adCmdTable      = 2 
    adCmdStoredProc = 4
    adCmdUnknown    = 8

    REM -- ADO cursor types
    adOpenForwardOnly = 0 '# (Default)
    adOpenKeyset      = 1
    adOpenDynamic     = 2
    adOpenStatic      = 3

    REM -- ADO lock types
    adLockReadOnly        = 1
    adLockPessimistic     = 2
    adLockOptimistic      = 3
    adLockBatchOptimistic = 4

    REM -- page constants
    top_margin   = "8"
    left_margin  = "8"
    color_link   = "#FF0000"
    color_alink  = "#FF0000"
    color_vlink  = "#FF0000"
    bgcolor      = "#FFFFFF"
    text         = "#000000"

    REM -- Create the Manager site dictionary
    Set mscsManagerSite = Server.CreateObject("Commerce.Dictionary")

    REM -- Get the dictionary for the site manager
    Set FD = Server.CreateObject("Commerce.FileDocument")
    path = Server.MapPath("/" + mscsPage.SiteRoot() + "/manager/config/site.csc")
    set mscsManagerSite = FD.ReadFromFile( path, "IISProperties")

    REM Set up links for other managers
    mscsShopperPath = "/" + mscsPage.SiteRoot() + "/manager/shopper-db/shopper_view.asp?shopper_id=:1"
    mscsReceiptsPath = "/" + mscsPage.SiteRoot() + "/manager/order-db/order_shopper_list.asp?shopper_id=:1"
    mscsBasketPath = "/" + mscsPage.SiteRoot() + "/manager/basket-db/basket_view.asp?shopper_id=:1"

    mscsProductPath = "/" + mscsPage.SiteRoot() + "/manager/catalog-vc/product_edit.asp?pfid=:1"
    mscsOrderPath = "/" + mscsPage.SiteRoot() + "/manager/orders-db/order_view.asp?order_id=:1"

    REM -- setup ADO connection
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.ConnectionTimeout = 15
    conn.CommandTimeout = 30
    conn.Open mscsManagerSite.DefaultConnectionString
    Set cmdTemp = Server.CreateObject("ADODB.Command")
    cmdTemp.CommandType = adCmdText
    Set cmdTemp.ActiveConnection = conn

    mscsTablePrefix = "vcturbo_"

%>

