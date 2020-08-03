/****************************************************************************
 *                                                                          *
 *   SCHEMA.SQL                                                             *
 *   Microsoft Commerce Server v3.00                                        *
 *                                                                          *
 *   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     *
 *                                                                          *
 ****************************************************************************/
/*
**  SQL server schema for VC TURBO STORE store with sub departments
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

create table vcturbo_dept (
            dept_id         int primary key not null,
            parent_id       int                 null,
            name            varchar(255)    not null,
            description     varchar(2000)       null,
            date_changed    date                null
            );

create table vcturbo_product_family (
            pfid            varchar(30) primary key not null,
            dept_id         int             not null,
            manufacturer_id int                 null,
            name            varchar(255)        not null,
            short_description   varchar(255)    null,
            long_description    varchar(2000)   null, 
            image_filename  varchar(255)        null,
            intro_date      date        null,
            date_changed    date            null,
            list_price      int     not null,
            monogramable    int         null
            );

create table vcturbo_product_variant (
            sku         int primary key not null,
            pfid        varchar(30) not null,
            attribute0  int         null,
            attribute1  int         null,
            attribute2  int         null,
            attribute3  int         null,
            attribute4  int         null
            );

create table vcturbo_product_attribute (
            pfid            varchar(30) not null,
            attribute_id    int         not null, /* 0-4, corresponds to Attribute0-Attribute3 */
            attribute_index int             null, /* 0: Attribute Name, >0: Attribute Value Indexes */
            attribute_value varchar(20) not null, /* Attribute value for a index */
            PRIMARY KEY (pfid, attribute_id, attribute_value)
            );

create table vcturbo_promo_upsell (
            pfid            varchar(30) not null,
            related_pfid    varchar(30) not null,
            description     varchar(255)    null,
            PRIMARY KEY(pfid, related_pfid)
            );

create table vcturbo_promo_cross (
            pfid            varchar(30)     not null,
            related_pfid    varchar(30)     not null,
        description varchar(255)    null,
        PRIMARY KEY(pfid, related_pfid)
            );

create table vcturbo_promo_price (
            promo_name       varchar(255) primary key not null,
            promo_type       int        not null,
            promo_description varchar(2000) null,
            promo_rank       int            null,
            active           int            null,
            date_start       date           null,
            date_end         date           null,
            shopper_all      int            null,
            shopper_column   varchar(64)    null,
            shopper_op       varchar(2)     null,
            shopper_value    varchar(64)    null,
            cond_all         int            null,
            cond_column      varchar(64)    null,
            cond_op          varchar(2)     null,
            cond_value       varchar(64)    null,
            cond_basis       char(1)        null,
            cond_min         int            null,
            award_all        int            null,
            award_column     varchar(64)    null,
            award_op         varchar(2)     null,
            award_value      varchar(64)    null,
            award_max        int            null,
        disjoint_cond_award int     null,
            disc_type        char(1)        null,
            disc_value       real           null
            );

create table vcturbo_shopper (
        shopper_id  char(32) primary key not null,
        created     date             null,
        name        varchar(235)         not null,
        password    varchar(20)      null,
        street      varchar(50)      null,
        city        varchar(50)      null,
        state       varchar(30)      null,
        zip     varchar(15)      null,
        country     varchar(20)      null,
        phone       varchar(16)      null,
        email       varchar(50)  unique  not null
        );

create table vcturbo_receipt_item (
        pfid            varchar(30) not null,
        sku             int         not null,
        order_id        char(26)    not null,
        row_id          int         not null,
        quantity        int             null,
        adjusted_price  int             null,
        PRIMARY KEY (order_id, row_id)
        );

create table vcturbo_receipt (
        order_id        char(26) primary key not null,
        shopper_id      char(32)        not null,
        total           int                 null,
        status          int                 null,
        date_entered    date                null,
        date_changed    date                null,
        marshalled_receipt  long raw        null
        );

create table vcturbo_basket (
        shopper_id      char(32) primary key,
        date_changed    date                null,
        marshalled_order long raw           null
        );

