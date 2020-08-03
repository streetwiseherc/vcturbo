/****************************************************************************
 *                                                                          *
 *   UNINSTALL.SQL                                                          *
 *   Microsoft Commerce Server v3.00                                        *
 *                                                                          *
 *   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     *
 *                                                                          *
 ****************************************************************************/
/*
  SQL server schema uninstall for VC TURBO STORE store with sub departments
*/
drop table vcturbo_receipt
go
drop table vcturbo_receipt_item
go
drop table vcturbo_shopper
go
drop table vcturbo_promo_upsell
go
drop table vcturbo_promo_cross
go
drop table vcturbo_promo_price
go
drop table vcturbo_product_attribute
go
drop table vcturbo_product_variant
go
drop table vcturbo_product_family
go
drop table vcturbo_dept
go
drop table vcturbo_basket
go
