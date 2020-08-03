/****************************************************************************
 *                                                                          *
 *   UNINSTALL.SQL                                                          *
 *   Microsoft Commerce Server v3.00                                        *
 *                                                                          *
 *   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     *
 *                                                                          *
 ****************************************************************************/
/*
**  SQL server schema uninstall for VC TURBO STORE store with sub departments
*/
/* drop existing tables: */
drop table vcturbo_receipt;
drop table vcturbo_receipt_item;
drop table vcturbo_shopper;
drop table vcturbo_promo_upsell;
drop table vcturbo_promo_cross;
drop table vcturbo_promo_price;
drop table vcturbo_product_attribute;
drop table vcturbo_product_variant;
drop table vcturbo_product_family;
drop table vcturbo_dept;
drop table vcturbo_basket;
