/****************************************************************************
 *                                                                          *
 *   LOADDATA.SQL                                                           *
 *   Microsoft Commerce Server v3.00                                        *
 *                                                                          *
 *   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     *
 *                                                                          *
 ****************************************************************************/

truncate table vcturbo_dept
go
insert into vcturbo_dept values (1,10, 'Gourmet Coffee', '','19960726 13:50')
go
insert into vcturbo_dept values (2,10,'Flavored Coffee','','19960726 13:50')
go
insert into vcturbo_dept values (3,10,'Varietals','','19960726 13:50')
go
insert into vcturbo_dept values (4,20,'Mugs','','19960726 13:50')
go
insert into vcturbo_dept values (5,20,'Grinders & Other Accessories','','19960726 13:50')
go
insert into vcturbo_dept values (6,20,'Brewers, Pots & Kettles','','19960726 13:50')
go
insert into vcturbo_dept values (7,30,'Clothing','','19960726 13:50')
go
insert into vcturbo_dept values (8,30,'Food Items','','19960726 13:50')
go
insert into vcturbo_dept values (10,0,'Gourmet Coffee','','19960807 13:17')
go
insert into vcturbo_dept values (20,0,'Brewing Accessories','','19960807 13:17')
go
insert into vcturbo_dept values (30,0,'Specialty Gifts','','19960807 13:17')
go


truncate table vcturbo_product_family
go
insert into vcturbo_product_family values ('1',1,1,'Ethiopian Supremo','This rare, West African blend comes from the ''birth place'' of coffee.  Still harvested from wild trees,  Ethiopian Supremo is famous for its light body, full flavor and smoothness.  The delicate fruity flavor makes it great as a dessert coffee.', '','item1.gif','19960101','19960726 13:50',1195,0)
go
insert into vcturbo_product_family values ('10',2,1,'Chocolate Noisette', 'We took the enticing combination of  Frangelico, creme and coffee one step further by introducing a light dusting of chocolate throughout.  This unique blending creates a magnificent coffee perfect for any time of the day.', '','item10.gif','19960101','19960726 13:50',1095,0)
go
insert into vcturbo_product_family values ('11',2,1,'Shamrock Creme Truffle', 'This scrumptious coffee boasts a richness of an Irish Creme liqueur flavor and a trace of sinfully sweet chocolate.  It''s a rich and rewarding experience.', '','item11.gif','19960101','19960726 13:50',795,0)
go
insert into vcturbo_product_family values ('12',2,1,'Irish Creme', 'The delicious combination of coffee with the delectable taste of Irish Creme liqueur creates one of the most popular coffees around.  Top with whipped cream for an exceptionally grand dessert all on its own!', '','item12.gif','19960101','19960726 13:50',1095,0)
go
insert into vcturbo_product_family values ('14',4,1,'Plastic Commuter Mug', 'For those of you who like the idea of a traveling container, you''ll love our generously sized 16-oz. mug. Along with the splash-proof lid and double-wall insulation, this innovative design features a no-skid base.', '','item14.gif','19960101','19960726 13:50',995,0)
go
insert into vcturbo_product_family values ('17',6,1,'French Press', 'The most elegant way to serve your guests at the table.  Having made its way from Italy to France and, now, with its growing popularity in the U.S., the French Press is one of the easiest brewing machines around.', 'With its easy-to-grasp handle and low-profile lid, this design is sure to become your new favorite.  And just wait until you taste the full-flavored coffee you''ve just brewed!','item17.gif','19960101','19960726 13:50',2195,0)
go
insert into vcturbo_product_family values ('18',6,1,'Automatic Drip', 'This superlative Auto Filter Drip coffee maker was designed to bring fresh water to the ideal temperature before ever reaching the coffee grounds.  This ensures the fullest flavor in every cup of coffee you brew.', 'With a removable, hinged water lid and filter, water level indicator and illuminated on/off switch, this is one coffee maker not to be passed up!','item18.gif','19960101','19960726 13:50',5495,0)
go
insert into vcturbo_product_family values ('19',6,1,'Potbellied Teapot', 'Our gracefully classic 32-oz. potbellied teapot is the perfect addition to any home.  Brew tea for yourself or for those special guests who drop by.  It will comfortably contain 2-3 tea bags to brew one perfect pot of tea.', '','item19.gif','19960101','19960726 13:50',995,0)
go
insert into vcturbo_product_family values ('2',1,1,'Majestic Mountain Blend', 'This full-bodied, aromatic coffee achieves a good balance of acidity and flavor  consistent to the final cup.  Clean, strong, and pleasing, with a sharpness that makes this blend perfect for any occasion.', '','item2.gif','19960101','19960726 13:50',995,0)
go
insert into vcturbo_product_family values ('20',6,1,'Small Espresso Machine', 'This magnificent machine offers more features than you would normally expect to find in this size and price range.  With its removable drip tray, powerful swiveling steam wand and compact size, what more could you want?', '','item20.gif','19960101','19960726 13:50',19995,0)
go
insert into vcturbo_product_family values ('21',5,1,'Steaming Pitcher', 'A must have for any espresso machine owner!  This generously sized stainless steel steaming pitcher has the capacity to hold up to 20 oz.  of milk for your favorite espresso drinks.  Whether you love Cappuccinos or Mochas you can''t do without this!', '','item21.gif','19960101','19960726 13:50',995,0)
go
insert into vcturbo_product_family values ('22',5,1,'Grinder', 'Why spend a fortune on a huge, cumbersome grinder when this smaller version has the ability to handle it all?  Whether it''s a coarse or cone-grind texture you''re after, this blade grinder spin beans throughout resulting in a perfect grind.', 'With a removable lid, the grinds fall effortlessly out of the holding container.  What could be easier?','item22.gif','19960101','19960726 13:50',1995,0)
go
insert into vcturbo_product_family values ('23',5,1,'Thermometer', 'Made to attach easily to the side of a steaming pitcher, this thermometer ensures you''ll steam your milk to the proper 165 degrees required to create a perfect-tasting espresso drink every time!', '','item23.gif','19960101','19960726 13:50',595,0)
go
insert into vcturbo_product_family values ('24',5,1,'Tamper', 'No espresso maker would be complete without one of these.  This hand-carved, wooden tamper uniformly packs grinds into the portafilter for the most amazing shot of espresso possible.', '','item24.gif','19960101','19960726 13:50',595,0)
go
insert into vcturbo_product_family values ('3',1,1,'Mezzana Blend', 'This distinctively sweet South American blend is known to have a velvety body with low acidity and rich flavors.  Served slightly on the weak side, its delicate flavors and aromas will satisfy even the most discerning palate.', '','item3.gif','19960101','19960726 13:50',995,0)
go
insert into vcturbo_product_family values ('31',8,1,'Chocolate Espresso Beans', 'Savor the richness of silky dark chocolate and enjoy the bold taste of French Roast Beans.  This marriage of chocolate and espresso bean has been an exquisite treat in years past and will continue to make the perfect present for years to come.', '','item31.gif','19960101','19960726 13:50',795,0)
go
insert into vcturbo_product_family values ('33',7,1,'Volcano T-shirt', 'Roomy, relaxed, and totally comfortable, our 100% cotton T-shirt features our logo screen printed on the front.  Wear this shirt and let everyone know where you get your favorite java!', '','item33.gif','19960101','19960726 13:50',1495,0)
go


insert into vcturbo_product_family values ('34',7,1,'Volcano Sweatshirt', 'Big, baggy, and completely cozy!  Our logo sweatshirt will be a great addition to anyone''s wardrobe.  Made of preshrunk cotton, it''s guaranteed to fit for life.', '','item34.gif','19960101','19960726 13:50',1995,0)
go
insert into vcturbo_product_family values ('35',7,1,'Button-Down Shirt', 'For those of you who want a uniquely individualized shirt for work or play, we have several styles of our button-down shirt to offer.  In an array of sizes, styles, and colors, there''s bound to be a shirt for you!', '','item35.gif','19960101','19960726 13:50',2495,1)
go
insert into vcturbo_product_family values ('36',7,1,'Baseball Jersey', 'This jersey is a great alternative for those who want a cross between the T-shirt and the Button-Down.  Made from 100% preshrunk cotton, this shirt features our logo on the front, upper-left corner and nine buttons from top to bottom.', '','item36.gif','19960101','19960726 13:50',1695,0)  
go
insert into vcturbo_product_family values ('4',1,1,'East Timor Estate', 'A rare washed Indonesian Arabica from East Timor, this heavily bodied blend has a hint of hidden spices and oils.  The cinnamon-scented aroma and rich flavors make this a pleasing after dinner coffee.', '','item4.gif','19960101','19960726 13:50',1295,0)   
go
insert into vcturbo_product_family values ('5',3,1,'Kona', 'Smooth, rich, and unusually aromatic, this coffee boasts unique character.  Our most popular coffee grown in the U.S. and with just one sip you''ll understand why.', '','item5.gif','19960101','19960726 13:50',995,0)         
go
insert into vcturbo_product_family values ('6',3,1,'Kenya', 'This delicate coffee from the highlands of East Africa features a winey, full-flavored body with an intriguing aroma.  Perfect for any time of the day, this is a must if you haven''t already savored the flavor!', '','item6.gif','19960101','19960726 13:50',895,0)                   
go
insert into vcturbo_product_family values ('7',3,1,'Colombian Supremo', 'This smooth, full-bodied coffee from Colombia features a sweet, delicate flavor, yet is extremely rich and aromatic with a pleasant sharpness and trace of winey overtones.  It is a classic coffee perfect for all occasions.', '','item7.gif','19960101','19960726 13:50',1095,0)                  
go
insert into vcturbo_product_family values ('8',3,1,'Costa Rica', 'This well-balanced coffee is known to be non-acidic, rich and round.  With its snappy, clean taste it has become a very distinguished coffee sure to please all coffee lovers!', '','item8.gif','19960101','19960726 13:50',895,0)                   
go
insert into vcturbo_product_family values ('9',2,1,'Hazelnut Creme', 'The sweet flavor of Frangelico and the subtle hint of creme are the perfect complements to the strong, Arabica coffee  they''re blended with.  Rich, soothing and flavorful, it''s the perfect coffee for any special occasion.', '','item9.gif','19960101','19960726 13:50',1095,0)                  
go



truncate table vcturbo_product_variant
go
insert into vcturbo_product_variant values (1,'12',0,0,0,0,0)
go
insert into vcturbo_product_variant values (3,'17',0,0,0,0,0)
go
insert into vcturbo_product_variant values (4,'18',0,0,0,0,0)
go
insert into vcturbo_product_variant values (5,'20',0,0,0,0,0)
go
insert into vcturbo_product_variant values (6,'21',0,0,0,0,0)
go
insert into vcturbo_product_variant values (7,'22',0,0,0,0,0)
go
insert into vcturbo_product_variant values (8,'23',0,0,0,0,0)
go
insert into vcturbo_product_variant values (9,'24',0,0,0,0,0)
go
insert into vcturbo_product_variant values (16,'31',0,0,0,0,0)
go
insert into vcturbo_product_variant values (28,'19',3,0,0,0,0)
go


insert into vcturbo_product_variant values (29,'19',1,0,0,0,0)
go
insert into vcturbo_product_variant values (30,'19',2,0,0,0,0)
go
insert into vcturbo_product_variant values (31,'36',3,2,0,0,0)
go
insert into vcturbo_product_variant values (32,'36',1,2,0,0,0)
go
insert into vcturbo_product_variant values (33,'36',2,2,0,0,0)
go
insert into vcturbo_product_variant values (34,'36',3,1,0,0,0)
go
insert into vcturbo_product_variant values (35,'36',1,1,0,0,0)
go
insert into vcturbo_product_variant values (36,'36',2,1,0,0,0)
go
insert into vcturbo_product_variant values (37,'34',1,1,0,0,0)
go
insert into vcturbo_product_variant values (38,'34',2,1,0,0,0)
go
insert into vcturbo_product_variant values (39,'34',3,1,0,0,0)
go
insert into vcturbo_product_variant values (40,'34',1,2,0,0,0)
go
insert into vcturbo_product_variant values (41,'34',2,2,0,0,0)
go
insert into vcturbo_product_variant values (42,'34',3,2,0,0,0)
go
insert into vcturbo_product_variant values (43,'33',2,3,0,0,0)
go
insert into vcturbo_product_variant values (44,'33',1,3,0,0,0)
go
insert into vcturbo_product_variant values (45,'33',2,2,0,0,0)
go
insert into vcturbo_product_variant values (46,'33',1,2,0,0,0)
go
insert into vcturbo_product_variant values (47,'33',2,1,0,0,0)
go
insert into vcturbo_product_variant values (48,'33',1,1,0,0,0)
go
insert into vcturbo_product_variant values (49,'14',1,2,0,0,0)
go
insert into vcturbo_product_variant values (50,'14',2,2,0,0,0)
go
insert into vcturbo_product_variant values (51,'14',3,2,0,0,0)
go
insert into vcturbo_product_variant values (52,'14',1,1,0,0,0)
go
insert into vcturbo_product_variant values (53,'14',2,1,0,0,0)
go
insert into vcturbo_product_variant values (54,'14',3,1,0,0,0)
go
insert into vcturbo_product_variant values (5041,'35',1,1,1,1,0)
go
insert into vcturbo_product_variant values (5042,'35',2,1,1,1,0)
go
insert into vcturbo_product_variant values (5043,'35',1,2,1,1,0)
go


insert into vcturbo_product_variant values (5044,'35',2,2,1,1,0)
go
insert into vcturbo_product_variant values (5045,'35',1,3,1,1,0)                                                                                                                                                                                                     
go
insert into vcturbo_product_variant values (5046,'35',2,3,1,1,0)           
go
insert into vcturbo_product_variant values (5047,'35',1,1,2,1,0)           
go
insert into vcturbo_product_variant values (5048,'35',2,1,2,1,0)           
go
insert into vcturbo_product_variant values (5049,'35',1,2,2,1,0)           
go
insert into vcturbo_product_variant values (5050,'35',2,2,2,1,0)           
go
insert into vcturbo_product_variant values (5051,'35',1,3,2,1,0)           
go
insert into vcturbo_product_variant values (5052,'35',2,3,2,1,0)
go
insert into vcturbo_product_variant values (5053,'35',1,1,3,1,0)
go
insert into vcturbo_product_variant values (5054,'35',2,1,3,1,0)
go
insert into vcturbo_product_variant values (5055,'35',1,2,3,1,0)
go
insert into vcturbo_product_variant values (5056,'35',2,2,3,1,0)
go
insert into vcturbo_product_variant values (5057,'35',1,3,3,1,0)
go
insert into vcturbo_product_variant values (5058,'35',2,3,3,1,0)
go
insert into vcturbo_product_variant values (5059,'35',1,1,4,1,0)
go
insert into vcturbo_product_variant values (5060,'35',2,1,4,1,0)
go
insert into vcturbo_product_variant values (5061,'35',1,2,4,1,0)
go
insert into vcturbo_product_variant values (5062,'35',2,2,4,1,0)
go
insert into vcturbo_product_variant values (5063,'35',1,3,4,1,0)
go
insert into vcturbo_product_variant values (5064,'35',2,3,4,1,0)
go
insert into vcturbo_product_variant values (5065,'35',1,1,1,2,0)
go
insert into vcturbo_product_variant values (5066,'35',2,1,1,2,0)
go
insert into vcturbo_product_variant values (5067,'35',1,2,1,2,0)
go
insert into vcturbo_product_variant values (5068,'35',2,2,1,2,0)
go
insert into vcturbo_product_variant values (5069,'35',1,3,1,2,0)
go
insert into vcturbo_product_variant values (5070,'35',2,3,1,2,0)
go
insert into vcturbo_product_variant values (5071,'35',1,1,2,2,0)
go
insert into vcturbo_product_variant values (5072,'35',2,1,2,2,0)
go
insert into vcturbo_product_variant values (5073,'35',1,2,2,2,0)
go


insert into vcturbo_product_variant values (5074,'35',2,2,2,2,0)
go
insert into vcturbo_product_variant values (5075,'35',1,3,2,2,0)
go
insert into vcturbo_product_variant values (5076,'35',2,3,2,2,0)
go
insert into vcturbo_product_variant values (5077,'35',1,1,3,2,0)
go
insert into vcturbo_product_variant values (5078,'35',2,1,3,2,0)
go
insert into vcturbo_product_variant values (5079,'35',1,2,3,2,0)
go
insert into vcturbo_product_variant values (5080,'35',2,2,3,2,0)
go
insert into vcturbo_product_variant values (5081,'35',1,3,3,2,0)
go
insert into vcturbo_product_variant values (5082,'35',2,3,3,2,0)
go
insert into vcturbo_product_variant values (5083,'35',1,1,4,2,0)
go
insert into vcturbo_product_variant values (5084,'35',2,1,4,2,0)
go
insert into vcturbo_product_variant values (5085,'35',1,2,4,2,0)
go
insert into vcturbo_product_variant values (5086,'35',2,2,4,2,0)
go
insert into vcturbo_product_variant values (5087,'35',1,3,4,2,0)
go
insert into vcturbo_product_variant values (5088,'35',2,3,4,2,0)
go
insert into vcturbo_product_variant values (5089,'35',1,1,1,3,0)
go
insert into vcturbo_product_variant values (5090,'35',2,1,1,3,0)
go
insert into vcturbo_product_variant values (5091,'35',1,2,1,3,0)
go
insert into vcturbo_product_variant values (5092,'35',2,2,1,3,0)
go
insert into vcturbo_product_variant values (5093,'35',1,3,1,3,0)
go
insert into vcturbo_product_variant values (5094,'35',2,3,1,3,0)
go
insert into vcturbo_product_variant values (5095,'35',1,1,2,3,0)
go
insert into vcturbo_product_variant values (5096,'35',2,1,2,3,0)
go
insert into vcturbo_product_variant values (5097,'35',1,2,2,3,0)
go
insert into vcturbo_product_variant values (5098,'35',2,2,2,3,0)
go
insert into vcturbo_product_variant values (5099,'35',1,3,2,3,0)
go
insert into vcturbo_product_variant values (5100,'35',2,3,2,3,0)
go
insert into vcturbo_product_variant values (5101,'35',1,1,3,3,0)                                                                                                                                                                                                     
go
insert into vcturbo_product_variant values (5102,'35',2,1,3,3,0)                                                                                                                                                                                                     
go
insert into vcturbo_product_variant values (5103,'35',1,2,3,3,0)                                                                                                                                                                                                     
go


insert into vcturbo_product_variant values (5104,'35',2,2,3,3,0)                                                                                                                                                                                                     
go
insert into vcturbo_product_variant values (5105,'35',1,3,3,3,0)                                                                                                                                                                                                     
go
insert into vcturbo_product_variant values (5106,'35',2,3,3,3,0)                                                                                                                                                                                                     
go
insert into vcturbo_product_variant values (5107,'35',1,1,4,3,0)                                                                                                                                                                                                     
go
insert into vcturbo_product_variant values (5108,'35',2,1,4,3,0)                                                                                                                                                                                                     
go
insert into vcturbo_product_variant values (5109,'35',1,2,4,3,0)                                                                                                                                                                                                     
go
insert into vcturbo_product_variant values (5110,'35',2,2,4,3,0)                                                                                                                                                                                                     
go
insert into vcturbo_product_variant values (5111,'35',1,3,4,3,0)                                                                                                                                                                                                     
go
insert into vcturbo_product_variant values (5112,'35',2,3,4,3,0)                                                                                                                                                                                                     
go
insert into vcturbo_product_variant values (6355,'1',0,0,0,0,0)
go
insert into vcturbo_product_variant values (6356,'2',0,0,0,0,0)
go
insert into vcturbo_product_variant values (6357,'3',0,0,0,0,0)
go
insert into vcturbo_product_variant values (6358,'4',0,0,0,0,0)
go
insert into vcturbo_product_variant values (6359,'5',0,0,0,0,0)
go
insert into vcturbo_product_variant values (6360,'6',0,0,0,0,0)
go
insert into vcturbo_product_variant values (6361,'7',0,0,0,0,0)
go
insert into vcturbo_product_variant values (6362,'8',0,0,0,0,0)
go
insert into vcturbo_product_variant values (6363,'9',0,0,0,0,0)
go
insert into vcturbo_product_variant values (6364,'10',0,0,0,0,0)
go
insert into vcturbo_product_variant values (6365,'11',0,0,0,0,0)
go


truncate table vcturbo_product_attribute
go
insert into vcturbo_product_attribute values ('35',2,2,'White')                                      
go
insert into vcturbo_product_attribute values ('35',2,3,'Denim')                                      
go
insert into vcturbo_product_attribute values ('35',2,4,'Khaki')                                      
go
insert into vcturbo_product_attribute values ('35',2,1,'Hunter')                                     
go
insert into vcturbo_product_attribute values ('36',1,2,'Blue')
go
insert into vcturbo_product_attribute values ('36',1,1,'Green')                                      
go
insert into vcturbo_product_attribute values ('34',1,1,'Orange')                                     
go
insert into vcturbo_product_attribute values ('34',1,2,'Purple')                                     
go
insert into vcturbo_product_attribute values ('33',0,1,'Black')                                      
go
insert into vcturbo_product_attribute values ('33',0,2,'Tan')                                        
go
insert into vcturbo_product_attribute values ('35',3,1,'Button Down Broad Cl')                       
go
insert into vcturbo_product_attribute values ('35',3,2,'Relaxed Button Down')                        
go
insert into vcturbo_product_attribute values ('35',3,3,'Regular Broad Cloth')
go
insert into vcturbo_product_attribute values ('35',0,1,'15')
go
insert into vcturbo_product_attribute values ('35',0,2,'16')                                         
go
insert into vcturbo_product_attribute values ('35',1,1,'32')                                         
go
insert into vcturbo_product_attribute values ('35',1,2,'33')                                         
go


insert into vcturbo_product_attribute values ('35',1,3,'34')                                         
go
insert into vcturbo_product_attribute values ('33',1,3,'Medium')                                     
go
insert into vcturbo_product_attribute values ('33',1,2,'Large')                                      
go
insert into vcturbo_product_attribute values ('33',1,1,'X-Large')                                    
go
insert into vcturbo_product_attribute values ('34',0,3,'Medium')                                     
go
insert into vcturbo_product_attribute values ('34',0,2,'Large')                                      
go
insert into vcturbo_product_attribute values ('34',0,1,'X-Large')                                    
go
insert into vcturbo_product_attribute values ('36',0,1,'Medium')                                     
go
insert into vcturbo_product_attribute values ('36',0,2,'Large')                                      
go
insert into vcturbo_product_attribute values ('36',0,3,'X-Large')                                    
go
insert into vcturbo_product_attribute values ('14',1,2,'Blue')                                       
go
insert into vcturbo_product_attribute values ('14',1,1,'Black')                                      
go
insert into vcturbo_product_attribute values ('14',0,3,'Short')                                      
go
insert into vcturbo_product_attribute values ('14',0,2,'Tall')                                       
go
insert into vcturbo_product_attribute values ('14',0,1,'Grande')                                     
go
insert into vcturbo_product_attribute values ('19',0,3,'Green')                                      
go
insert into vcturbo_product_attribute values ('19',0,1,'Blue')                                       
go
insert into vcturbo_product_attribute values ('19',0,2,'Purple')                                     
go
insert into vcturbo_product_attribute values ('14',1,0,'Color')                                      
go
insert into vcturbo_product_attribute values ('14',0,0,'Size')                                       
go
insert into vcturbo_product_attribute values ('19',0,0,'Color')                                      
go
insert into vcturbo_product_attribute values ('33',0,0,'Color')                                      
go
insert into vcturbo_product_attribute values ('33',1,0,'Size')                                       
go
insert into vcturbo_product_attribute values ('34',1,0,'Color')                                      
go


insert into vcturbo_product_attribute values ('34',0,0,'Size')                                       
go
insert into vcturbo_product_attribute values ('35',1,0,'Arm Length')                                 
go
insert into vcturbo_product_attribute values ('35',2,0,'Color')                                      
go
insert into vcturbo_product_attribute values ('35',0,0,'Neck Size')                                  
go
insert into vcturbo_product_attribute values ('35',3,0,'Style')                                      
go
insert into vcturbo_product_attribute values ('36',1,0,'Color')                                      
go
insert into vcturbo_product_attribute values ('36',0,0,'Size')                                       
go


truncate table vcturbo_promo_upsell
go
insert into vcturbo_promo_upsell values ('2','4','Why not get something better for a change?')
go
insert into vcturbo_promo_upsell values ('3','4','Why not get something better for a change?')
go
insert into vcturbo_promo_upsell values ('11','9','Why not get something better for a change?')
go
insert into vcturbo_promo_upsell values ('6','7','Why not get something better for a change?')
go
insert into vcturbo_promo_upsell values ('8','7','Why not get something better for a change?')
go
insert into vcturbo_promo_upsell values ('33','35','Why not get something better for a change?')
go

truncate table vcturbo_promo_cross
go
insert into vcturbo_promo_cross values ('3','12', '')
go
insert into vcturbo_promo_cross values ('24','22', '')
go
insert into vcturbo_promo_cross values ('33','36', '')
go

truncate table vcturbo_promo_price
go

insert into vcturbo_promo_price values ('Grinder and Ethiopian Coffee',100,'Buy a Grinder get Ethiopian Coffee at 40% off',10,1,'19970812','19990831',1,null,null,null,0,'pfid','=','22','Q',1,0,'pfid','=','1', 1, 1,'%',40.0)
go

