/****************************************************************************
 *                                                                          *
 *   SCHEMA.SQL                                                             *
 *   Microsoft Commerce Server v3.00                                        *
 *                                                                          *
 *   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     *
 *                                                                          *
 ****************************************************************************/
/*
  SQL server schema for VC TURBO STORE store with sub departments
*/
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
drop table vcturbo_basket_item
go
drop table vcturbo_receipt
go
drop table vcturbo_receipt_item
go
create table vcturbo_dept (
        dept_id     int primary key     not null,
        parent_id   int             null,
        name        varchar(255)        not null,
        description text            null, /* changed from varchar(2000), */
        date_changed    datetime        null
        )
go
create table vcturbo_product_family (
        pfid        varchar(30) primary key not null,
        dept_id     int         not null,
        manufacturer_id int         null,
        name        varchar(255)        not null,
        short_description   varchar(255)    null,
        long_description    text        null,
        image_filename  varchar(255)        null,
        intro_date  datetime        null,
        date_changed    datetime        null,
        list_price  int         not null,
        monogramable    tinyint         null
        )
go
create table vcturbo_product_variant (
        sku     int primary key not null,
        pfid        varchar(30) not null,
        attribute0  tinyint     null,
        attribute1  tinyint     null,
        attribute2  tinyint     null,
        attribute3  tinyint     null,
        attribute4  tinyint     null
        )
go
create table vcturbo_product_attribute (
        pfid        varchar(30) not null,
        attribute_id    tinyint not null,     /* 0-4, corresponds to Attribute0-Attribute3 */
        attribute_index tinyint     null,     /* 0: Attribute Name, >0: Attribute Value Indexes */
        attribute_value varchar(20) not null, /* Attribute value for a index */
    PRIMARY KEY (pfid, attribute_id, attribute_value)
        )
go
create table vcturbo_promo_upsell (
        pfid        varchar(30) not null,
        related_pfid    varchar(30) not null,
        description varchar(255)    null,
        PRIMARY KEY(pfid, related_pfid)
        )
go
create table vcturbo_promo_cross (
        pfid            varchar(30)         not null,
        related_pfid    varchar(30)         not null,
    description varchar(255)    null,
    PRIMARY KEY(pfid, related_pfid)
        )
go
create table vcturbo_promo_price ( 
        promo_name        varchar(255) primary key not null,
        promo_type        int       not null,
        promo_description text          null,
        promo_rank        int           null,
        active            int           null,
        date_start        datetime      null,
        date_end          datetime      null,
        shopper_all       int           null ,
        shopper_column    varchar(64)   null,
        shopper_op        varchar(2)    null,
        shopper_value     varchar(64)   null,
        cond_all          int           null,
        cond_column       varchar(64)   null,
        cond_op           varchar(2)    null,
        cond_value        varchar(64)   null,
        cond_basis        char(1)       null,
        cond_min          int           null, 
        award_all         int           null,
        award_column      varchar(64)   null,
        award_op          varchar(2)    null,
        award_value       varchar(64)   null,
        award_max         int           null,
        disjoint_cond_award int     null,
        disc_type         char(1)       null,
        disc_value        real          null
        )
go
create table vcturbo_shopper (
        shopper_id  char(32) primary key    not null,
        created     datetime            null,
        name        varchar(235)        not null,
        password    varchar(20)         null,
        street      varchar(50)         null,
        city        varchar(50)         null,
        state       varchar(30)         null,
        zip     varchar(15)         null,
        country     varchar(20)         null,
        phone       varchar(16)         null,
        email       varchar(50) unique  not null,
        )
go
create table vcturbo_receipt (
        order_id    		int primary key 	not null,
        
        shopper_id  		char(32) 			    not null,
        
        created     		datetime 				null,
        modified    		datetime 				null,
        
        ship_to_name        varchar(235)        	null,
        ship_to_street      varchar(50)         	null,
        ship_to_city        varchar(50)         	null,
        ship_to_state       varchar(30)         	null,
        ship_to_zip     	varchar(15)         	null,
        ship_to_country     varchar(20)         	null,
        ship_to_phone       varchar(16)         	null,
        ship_to_email       varchar(50) 		  	null,

        bill_to_name        varchar(235)        	null,
        bill_to_street      varchar(50)         	null,
        bill_to_city        varchar(50)         	null,
        bill_to_state       varchar(30)         	null,
        bill_to_zip     	varchar(15)         	null,
        bill_to_country     varchar(20)         	null,
        bill_to_phone       varchar(16)         	null,
        bill_to_email       varchar(50) 		  	null,

        oadjust_subtotal	int						null,
        shipping_total		int						null,
        handling_total		int						null,
        tax_total			int						null,
        tax_included		int						null,

        cc_name				varchar(235)			null,
        cc_type				varchar(50)				null,
        
        total_total			int						null
        )
go
create table vcturbo_receipt_item(
		order_id			int 				not null,
		shopper_id			char(32)			not null,
		
		created				datetime				null,
		modified			datetime				null,

		item_id				int						null,

		sku					int						null,
		pfid				varchar(30)				null,
		name				varchar(255)			null,
		quantity			int						null,

		product_list_price		int					null,
		iadjust_regularprice	int					null,
		iadjust_currentprice	int					null,
		oadjust_adjustedprice	int					null,

		monogram			varchar(3)				null,
		attr_text			text					null
		)        
go		
create table vcturbo_basket (
        shopper_id  		char(32) primary key    not null,
        
        order_id    		int identity			not null,
        
        created     		datetime 				null,
        modified    		datetime 				null,
        
        ship_to_name        varchar(235)        	null,
        ship_to_street      varchar(50)         	null,
        ship_to_city        varchar(50)         	null,
        ship_to_state       varchar(30)         	null,
        ship_to_zip     	varchar(15)         	null,
        ship_to_country     varchar(20)         	null,
        ship_to_phone       varchar(16)         	null,
        ship_to_email       varchar(50)			  	null,

        bill_to_name        varchar(235)        	null,
        bill_to_street      varchar(50)         	null,
        bill_to_city        varchar(50)         	null,
        bill_to_state       varchar(30)         	null,
        bill_to_zip     	varchar(15)         	null,
        bill_to_country     varchar(20)         	null,
        bill_to_phone       varchar(16)         	null,
        bill_to_email       varchar(50)				null,

        cc_name				varchar(235)				null,        
        )
go
create table vcturbo_basket_item(
		shopper_id			char(32) 				not null,
		
		created				datetime				null,
		modified			datetime				null,

		item_id				numeric identity		not null,

		sku					int						null,
		pfid				varchar(30)				null,
		quantity			int						null,
		placed_price		int						null,

		monogram			varchar(3)				null,
		attr_text			text					null,

		name				varchar(255)			null
		)

