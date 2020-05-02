import pandas as pd
import requests
from bs4 import BeautifulSoup as soup


def main():
    try:
        print("enter the Url.")
        url = str(input()).strip()
        # https://hipcitymerch.com/collections/all/products.json
        # https://hipcitymerch.com/collections/lifestyle/products.json
        # https://hipcitymerch.com/collections/coding/products.json
        # https://hipcitymerch.com/collections/nerd/products.json
        # https://hipcitymerch.com/collections/lgbt/products.json
        # https://hipcitymerch.com/collections/gamer/products.json
        try:
            url_split = url.replace(".json", "").split("/")
            file_name = str(url_split[-2]) + "_" + str(url_split[-1]) + '_file.xlsx'

            req = requests.get(url)
            json_file = req.json()

            df_1_col = ["TemplateType=fptcustomcustom", "Version=2017.0829",
                        "The top 3 rows are for Amazon.com use only. Do not modify or delete the top 3 rows.", "", "",
                        "", "", "",
                        "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                        "Variation-Populate these attributes if your product is available in different variations (for example color or wattage)",
                        "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]

            df_2_col = ["Seller SKU", "Product ID", "Product ID Type", "Product Name", "Outer Material Type",
                        "Outer Material Type",
                        "Material Composition", "Material Composition", "Product Description", "Product Type",
                        "Department", "Brand Name",
                        "Item Type Keyword", "Update Delete", "Standard Price", "Sale Price", "Sale Start Date",
                        "Sale End Date", "Quantity",
                        "Shipping-Template", "Handling Time", "Shipping Weight",
                        "Website Shipping Weight Unit Of Measure", "Key Product Features1",
                        "Key Product Features2", "Key Product Features3", "Key Product Features4",
                        "Key Product Features5", "Main Image URL",
                        "Other Image URL1", "Parentage", "Parent SKU", "Relationship Type", "Variation Theme", "Size",
                        "Size Map", "Color",
                        "Color Map", "Search Terms", "Product Tax Code", "Fit Type", "NeckStyle", "Pattern Style",
                        "Style", "theme",
                        "Top Style", "Sleeve Type", "Strap Type", "Closure Type", "Outer Material Type",
                        "Material Composition",
                        "Other Image URL3", "Other Image URL4"]

            df_3_col = ["item_sku", "external_product_id", "external_product_id_type", "item_name",
                        "outer_material_type1",
                        "outer_material_type2", "material_composition1", "material_composition2", "product_description",
                        "feed_product_type",
                        "department_name", "brand_name", "item_type", "update_delete", "standard_price", "sale_price",
                        "sale_from_date",
                        "sale_end_date", "quantity", "merchant_shipping_group_name", "fulfillment_latency",
                        "website_shipping_weight",
                        "website_shipping_weight_unit_of_measure", "bullet_point1", "bullet_point2", "bullet_point3",
                        "bullet_point4",
                        "bullet_point5", "main_image_url", "other_image_url1", "parent_child", "parent_sku",
                        "relationship_type",
                        "variation_theme", "size_name", "size_map", "color_name", "color_map", "generic_keywords",
                        "product_tax_code", "fit_type",
                        "neck_style", "pattern_type", "style_name", "theme", "top_style", "sleeve_type", "strap_type",
                        "closure_type",
                        "outer_material_type3", "material_composition3", "other_image_url3", "other_image_url4"]

            writer = pd.ExcelWriter(file_name, engine='openpyxl')

            dfObj = pd.DataFrame(columns=df_1_col)
            dfObj_1 = pd.DataFrame(columns=df_2_col)
            dfObj_2 = pd.DataFrame(columns=df_3_col)

            dfObj.to_excel(writer, index=False)
            dfObj_1.to_excel(writer, startrow=len(dfObj) + 1, index=False)
            dfObj_2.to_excel(writer, startrow=len(dfObj) + len(dfObj_1) + 2, index=False)

            writer.save()

            read_file = json_file
            product_all_lst = read_file["products"]

            try:
                bullet_point2_url = "https://hipcitymerch.com/collections/" + str(url_split[-2]) + "/"
                bullet_req = requests.get(bullet_point2_url)
                page_soup = soup(bullet_req.content, "html5lib")

                bullet_point2_txt = page_soup.find("div", {"class": "rte rte--header"}).text.strip()
            except:
                bullet_point2_txt = ""
                pass

            item_sku_value = 11000
            for i, product in enumerate(product_all_lst):
                try:
                    item_sku, external_product_id, external_product_id_type, item_name, outer_material_type1 = '', '', '', '', ''
                    outer_material_type2, material_composition1, material_composition2, product_description = '', '', '', ''
                    feed_product_type, department_name, brand_name, item_type, update_delete, standard_price = '', '', '', '', '', ''
                    sale_price, sale_from_date, sale_end_date, quantity, merchant_shipping_group_name = '', '', '', '', ''
                    fulfillment_latency, website_shipping_weight, website_shipping_weight_unit_of_measure = '', '', ''
                    bullet_point1, bullet_point2, bullet_point3, bullet_point4, bullet_point5, main_image_url = '', '', '', '', '', ''
                    other_image_url1, parent_child, parent_sku, relationship_type, variation_theme, size_name = '', '', '', '', '', ''
                    size_map, color_name, color_map, generic_keywords, product_tax_code, fit_type, neck_style = '', '', '', '', '', '', ''
                    pattern_type, style_name, theme, top_style, sleeve_type, strap_type, closure_type = '', '', '', '', '', '', ''
                    outer_material_type3, material_composition3, other_image_url3, other_image_url4 = '', '', '', ''

                    brand_name = product["vendor"]
                    product_title_name = product["title"]
                    item_name = product_title_name

                    if (
                            item_name.lower().__contains__(("T-Shirt").lower()) or
                            item_name.lower().__contains__(("Tank Top").lower())
                    ):
                        feed_product_type = "shirt"
                    elif item_name.lower().__contains__(("Joggers").lower()):
                        feed_product_type = "pants"
                    elif (
                            item_name.lower().__contains__(("Sweatshirt").lower()) or
                            item_name.lower().__contains__(("Hoodie").lower())
                    ):
                        feed_product_type = "sweater"
                    else:
                        feed_product_type = ''

                    external_product_id = ''

                    x = brand_name.split(" ")
                    sk = ''
                    for index, item in enumerate(x):
                        if index == 0:
                            sk = sk + item[0].upper()
                        else:
                            sk = sk + item[0].lower()

                    color_option_1 = str(product["variants"][0]["option1"]).strip().upper()
                    color_option_2 = str(product["variants"][0]["option2"]).strip().upper()

                    if color_option_1 == "NONE" or color_option_2 == "NONE":
                        color_sku = 'Black'
                    elif len(color_option_1) > 3:
                        color_sku = color_option_1
                    else:
                        color_sku = color_option_2

                    product_item_sku = sk + "-" + str(item_sku_value)
                    item_sku = sk +"-"+str(item_sku_value) + "-" + str(color_sku).upper()

                    if item_sku == "Hcm11016-ATHLETIC HEATHER":
                        de = ''

                    external_product_id_type = ''

                    if ("T-Shirt").lower() in item_name.lower():
                        item_type = "novelty-t-shirts"
                    elif ("Women's Racerback Tank Top, Slim").lower() in item_name.lower():
                        item_type = "novelty-tank-tops"
                    elif ("Joggers").lower() in item_name.lower():
                        item_type = "athletic-pants"
                    elif ("Sweatshirt").lower() in item_name.lower():
                        item_type = "fashion-sweatshirts"
                    elif ("Hoodie").lower() in item_name.lower():
                        item_type = "fashion-hoodies"
                    else:
                        item_type = ''

                    product_description = str(product["body_html"]).replace("<br>", "").replace("</br>", "").strip()
                    product_description = str(product_description).replace("<p>", "").replace("</p>", "").strip()
                    product_description = str(product_description).replace("<div>", "").replace("</div>", "").strip()
                    product_description = str(product_description).replace("<span>", "").replace("</span>", "").strip()

                    bullet_point1 = product_description
                    bullet_point2 = bullet_point2_txt

                    if "Women's Racerback Tank Top".lower() in item_name.lower():
                        bullet_point3 = "High quality print of this slim fit tank-top will turn heads. And bystanders won't be " \
                                        "disappointed - the racerback cut looks good one any woman's shoulders."
                        bullet_point4 = "Made with 3.9oz 60/40 cotton poly, the 30 single lightweight jersey drapes nicely " \
                                        "against the body. The shirts are pre-laundered to reduce shrinkage, and it's made " \
                                        "with self fabric on the neck and armhole binding."
                        bullet_point5 = "This fine gauge blended fabric represents a new benchmark in value. You will be amazed " \
                                        "by the ultra-soft hand and performance of this exclusive fabric. Lightweight combed " \
                                        "ring-spun and poly blend. Racerback. Colors are solid unless otherwise noted."
                        other_image_url1 = "https://cdn.shopify.com/s/files/1/0188/8462/4435/files/Next_level_Women_s_Ideal_Racerback_Tank.jpg?v=1587633898"

                    elif "Women's T-Shirt".lower() in item_name.lower():
                        bullet_point3 = "Our t-shirt is definitely worthy of becoming your new fave. Comfortable, form-fitting, " \
                                        "and a stylish base for any outfit. This shirt tends to run small – order a size up if " \
                                        "you want a looser fit."
                        bullet_point4 = "Her go-to tee fits like a well-loved favorite, featuring a slim feminine fit. " \
                                        "Additionally, it is really comfortable - an item to fall in love with."
                        bullet_point5 = "Our Ladies' t-shirt is made of 100% combed ringspun cotton, 30 singles. It has a " \
                                        "longer length body and shoulder taping with a tear away label."
                        other_image_url1 = "https://cdn.shopify.com/s/files/1/0188/8462/4435/files/BELLA_CANVAS_WOMEN_S_FAVORITE_TEE.jpg?v=1587119289"

                    elif "Men's T-Shirt".lower() in item_name.lower():
                        bullet_point3 = "Our t-shirt is definitely worthy of becoming your new fave. Comfortable, form-fitting, " \
                                        "and a stylish base for any outfit. This shirt tends to run small – order a size up if " \
                                        "you want a looser fit."
                        bullet_point4 = "Our softstyle t-shirt made of 100% softstyle cotton and 30 singles. It as double-needle" \
                                        " stitched neckline and sleeves and tear away label. It is quarter-turned, taped neck" \
                                        " and shoulder with a three-quarter inch seamless collar."
                        bullet_point5 = "a great staple t-shirt that compliments any outfit. It's made of a heavier " \
                                        "cotton and the double-stitched neckline and sleeves give it more durability, " \
                                        "so it can become an everyday favorite."
                        other_image_url1 = "https://cdn.shopify.com/s/files/1/0188/8462/4435/files/Gildan_Men.jpg?v=1587010913"

                    elif "Women's Joggers".lower() in item_name.lower():
                        bullet_point3 = "They're made from breathable and soft cotton and polyester blend—just what's needed " \
                                        "for a great workout, or a quiet evening at home."
                        bullet_point4 = "Soft cotton-feel fabric face with a Brushed fleece fabric inside. Our joggers " \
                                        "have practical pockets with an elastic waistband. A white drawstring is also " \
                                        "included for an easier fit."
                        bullet_point5 = " "
                        other_image_url1 = "https://cdn.shopify.com/s/files/1/0188/8462/4435/files/JOGGERS-1024.jpg?v=1588305186"

                    elif "Men's Joggers".lower() in item_name.lower():
                        bullet_point3 = "They're made from breathable and soft cotton and polyester blend—just what's needed " \
                                        "for a great workout, or a quiet evening at home."
                        bullet_point4 = "Soft cotton-feel fabric face with a Brushed fleece fabric inside. Our joggers " \
                                        "have practical pockets with an elastic waistband. A white drawstring is also " \
                                        "included for an easier fit."
                        bullet_point5 = " "
                        other_image_url1 = "https://cdn.shopify.com/s/files/1/0188/8462/4435/files/JOGGERS-1024.jpg?v=1588305186"

                    elif "Unisex-Adult Sweatshirt".lower() in item_name.lower():
                        bullet_point3 = "This soft sweatshirt has a loose fit for a comfortable feel. With durable print, it " \
                                        "will be a walking billboard for years to come."
                        bullet_point4 = "This well-loved Unisex Sweatshirt is the perfect addition to any wardrobe. " \
                                        "It has a crew neck, and it's made from air-jet spun yarn and quarter-turned fabric," \
                                        " which eliminates a center crease, reduces pilling, and gives the sweatshirt a soft," \
                                        " comfortable feel."
                        bullet_point5 = "Our crewneck sweatshirt made of a 50/50 blend of cotton and polyester. It has reduced " \
                                        "pilling and softer air-jet spun yarn with a 1x1 athletic rib knit collar, cuffs and " \
                                        "waistband, with spandex. The sweatshirt also features double-needle stitched collar, " \
                                        "shoulders, armholes, cuffs and waistband."
                        other_image_url1 = "https://cdn.shopify.com/s/files/1/0188/8462/4435/files/Gildan_Unisex_Sweatshirts_Size_Chart.jpg?v=1587010913"

                    elif "Unisex-Adult Hoodie".lower() in item_name.lower():
                        bullet_point3 = "Crafted for comfort, this lighter weight sweatshirt is perfect for relaxing. " \
                                        "Once put on, it will be impossible to take off."
                        bullet_point4 = "With a large front pouch pocket and drawstrings in a matching color, this Unisex Hoodie " \
                                        "is a sure crowd-favorite. It’s soft, stylish, and perfect for the cooler evenings."
                        bullet_point5 = "Our heavy blend hooded sweatshirt made of 50/50 cotton and polyester. This sweater " \
                                        "has a double-lined hood with matching drawcord (adult style only), reduced pilling " \
                                        "and softer air-jet spun yarn and 1x1 athletic rib knit cuffs and waistband with spandex. " \
                                        "It also has a front pouch pocket and has double-needle stitching throughout. Satin label."
                        other_image_url1 = "https://cdn.shopify.com/s/files/1/0188/8462/4435/files/Gildan_Unisex_Hoodies_Size_Chart.jpg?v=1587010913"
                    else:
                        bullet_point3 = ""
                        bullet_point4 = ""
                        bullet_point5 = ""
                        other_image_url1 = ""

                    if ("women").lower() in item_name.lower():
                        department_name = "womens"
                    elif ("men").lower() in item_name.lower():
                        department_name = "mens"
                    elif ("unisex").lower() in item_name.lower():
                        department_name = "unisex-adult"
                    else:
                        department_name = ''

                    main_image_url = product["variants"][0]['featured_image']['src']

                    other_image_url3 = ''
                    other_image_url4 = ''

                    parent_child = 'parent'
                    variation_theme = "sizecolor"

                    dfObj_2 = dfObj_2.append(
                        {"item_sku": item_sku, "external_product_id": external_product_id, "external_product_id_type":
                            external_product_id_type, "item_name": item_name,
                         "outer_material_type1": outer_material_type1,
                         "outer_material_type2": outer_material_type2, "material_composition1": material_composition1,
                         "material_composition2": material_composition2, "product_description": product_description,
                         "feed_product_type": feed_product_type, "department_name": department_name,
                         "brand_name": brand_name,
                         "item_type": item_type, "update_delete": update_delete, "standard_price": standard_price,
                         "sale_price":
                             sale_price, "sale_from_date": sale_from_date, "sale_end_date": sale_end_date,
                         "quantity": quantity,
                         "merchant_shipping_group_name": merchant_shipping_group_name,
                         "fulfillment_latency": fulfillment_latency,
                         "website_shipping_weight": website_shipping_weight, "website_shipping_weight_unit_of_measure":
                             website_shipping_weight_unit_of_measure, "bullet_point1": bullet_point1,
                         "bullet_point2": bullet_point2,
                         "bullet_point3": bullet_point3, "bullet_point4": bullet_point4, "bullet_point5": bullet_point5,
                         "main_image_url": main_image_url, "other_image_url1": other_image_url1,
                         "parent_child": parent_child,
                         "parent_sku": parent_sku, "relationship_type": relationship_type,
                         "variation_theme": variation_theme,
                         "size_name": size_name, "size_map": size_map, "color_name": color_name, "color_map": color_map,
                         "generic_keywords": generic_keywords, "product_tax_code": product_tax_code,
                         "fit_type": fit_type,
                         "neck_style": neck_style, "pattern_type": pattern_type, "style_name": style_name,
                         "theme": theme,
                         "top_style": top_style, "sleeve_type": sleeve_type, "strap_type": strap_type,
                         "closure_type": closure_type, "outer_material_type3": outer_material_type3,
                         "material_composition3":
                             material_composition3, "other_image_url3": other_image_url3,
                         "other_image_url4": other_image_url4},
                        ignore_index=True)

                    dfObj_2.to_excel(writer, startrow=len(dfObj) + len(dfObj_1) + 2, index=False)
                    writer.save()
                    item_sku_value += 1

                    if (
                            item_name.lower().__contains__("Women's T-Shirt, Slim".lower()) or
                            item_name.lower().__contains__("Women's T-Shirt".lower()) or
                            item_name.lower().__contains__("Men's T-Shirt, Regular".lower()) or
                            item_name.lower().__contains__("Men's T-Shirt".lower()) or
                            item_name.lower().__contains__("Women's Racerback Tank Top, Slim".lower()) or
                            item_name.lower().__contains__("Sweatshirt".lower()) or
                            item_name.lower().__contains__("Hoodie".lower())
                    ):
                        fulfillment_latency = "12"
                    elif "Joggers".lower() in item_name.lower():
                        fulfillment_latency = "20"
                    else:
                        fulfillment_latency = ""

                    quantity = '100'

                    fit_type = item_name.split(":")[-2].strip().split(" ")[-2]

                    neck_style = "crewneck"

                    pattern_type = ''
                    style_name = ''
                    generic_keywords = ''

                    theme = "humorous"

                    top_style = ''

                    sleeve_type = "short sleeve"

                    strap_type = ''
                    sale_price = ''
                    sale_from_date = ''
                    sale_end_date = ''
                    parent_child = 'child'
                    relationship_type = 'variation'

                    variants_lst = product["variants"]
                    for j, variant in enumerate(variants_lst):
                        try:
                            color_option_1 = str(variant["option1"]).strip().upper()
                            color_option_2 = str(variant["option2"]).strip().upper()

                            if color_option_1 == "NONE" or color_option_2 == "NONE":
                                color_name = color_sku
                            else:
                                color_name = str(variant["option1"]).strip()
                                if len(color_name) > 3:
                                    pass
                                else:
                                    color_name = str(variant["option2"]).strip()

                            color_map = color_name

                            size_name = str(variant["title"]).split("/")[-1].strip().upper()
                            size_map = size_name

                            item_sku = sk +"-"+ str(item_sku_value) + "-" + color_name.upper() + " / " + size_name

                            varient_title_item = product_title_name + " - " + color_name + " / " + size_name
                            item_name = varient_title_item

                            if (
                                    item_name.lower().__contains__("Women's T-Shirt, Slim".lower()) or
                                    item_name.lower().__contains__("Men's T-Shirt, Regular".lower()) or
                                    item_name.lower().__contains__("Women's Racerback Tank Top, Slim".lower()) or
                                    item_name.lower().__contains__("Sweatshirt".lower()) or
                                    item_name.lower().__contains__("Hoodie".lower())
                            ):
                                outer_material_type1 = "Cotton"
                            elif ("Joggers").lower() in item_name.lower():
                                outer_material_type1 = "Polyester"
                            else:
                                outer_material_type1 = ''

                            material_composition1 = outer_material_type1

                            if (
                                    item_name.lower().__contains__("Women's T-Shirt, Slim".lower()) or
                                    item_name.lower().__contains__("Men's T-Shirt, Regular".lower())
                            ):
                                outer_material_type2 = " "
                            elif (
                                    item_name.lower().__contains__("Women's Racerback Tank Top, Slim".lower()) or
                                    item_name.lower().__contains__("Sweatshirt".lower()) or
                                    item_name.lower().__contains__("Hoodie".lower())
                            ):
                                outer_material_type2 = "Polyester"
                            elif ("Joggers").lower() in item_name.lower():
                                outer_material_type2 = "Cotton"
                            else:
                                outer_material_type2 = ''

                            material_composition2 = outer_material_type2

                            if (
                                    item_name.lower().__contains__("Women's T-Shirt, Slim".lower()) or
                                    item_name.lower().__contains__("Men's T-Shirt, Regular".lower()) or
                                    item_name.lower().__contains__("Women's Racerback Tank Top, Slim".lower()) or
                                    item_name.lower().__contains__("Sweatshirt".lower()) or
                                    item_name.lower().__contains__("Hoodie".lower())
                            ):
                                outer_material_type3 = " "
                            elif "Hoodie".lower() in item_name.lower():
                                outer_material_type3 = "Elastane"
                            else:
                                outer_material_type3 = ''

                            material_composition3 = outer_material_type3

                            if "Women's T-Shirt, Slim".lower() in item_name.lower():
                                standard_price = str(18.99)
                            elif "Men's T-Shirt, Regular".lower() in item_name.lower():
                                standard_price = str(16.99)
                            elif "Women's Racerback Tank Top, Slim".lower() in item_name.lower():
                                standard_price = str(18.99)
                            elif "Joggers".lower() in item_name.lower():
                                standard_price = str(54.99)
                            elif "Sweatshirt".lower() in item_name.lower():
                                standard_price = str(31.99)
                            elif "Hoodie".lower() in item_name.lower():
                                standard_price = str(37.99)
                            else:
                                standard_price = ""

                            main_image_url = variant['featured_image']['src']

                            parent_sku = product_item_sku
                            update_delete = ''

                            dfObj_2 = dfObj_2.append(
                                {"item_sku": item_sku, "external_product_id": external_product_id,
                                 "external_product_id_type":
                                     external_product_id_type, "item_name": item_name,
                                 "outer_material_type1": outer_material_type1,
                                 "outer_material_type2": outer_material_type2,
                                 "material_composition1": material_composition1,
                                 "material_composition2": material_composition2,
                                 "product_description": product_description,
                                 "feed_product_type": feed_product_type, "department_name": department_name,
                                 "brand_name": brand_name,
                                 "item_type": item_type, "update_delete": update_delete,
                                 "standard_price": standard_price, "sale_price":
                                     sale_price, "sale_from_date": sale_from_date, "sale_end_date": sale_end_date,
                                 "quantity": quantity,
                                 "merchant_shipping_group_name": merchant_shipping_group_name,
                                 "fulfillment_latency": fulfillment_latency,
                                 "website_shipping_weight": website_shipping_weight,
                                 "website_shipping_weight_unit_of_measure":
                                     website_shipping_weight_unit_of_measure, "bullet_point1": bullet_point1,
                                 "bullet_point2": bullet_point2,
                                 "bullet_point3": bullet_point3, "bullet_point4": bullet_point4,
                                 "bullet_point5": bullet_point5,
                                 "main_image_url": main_image_url, "other_image_url1": other_image_url1,
                                 "parent_child": parent_child,
                                 "parent_sku": parent_sku, "relationship_type": relationship_type,
                                 "variation_theme": variation_theme,
                                 "size_name": size_name, "size_map": size_map, "color_name": color_name,
                                 "color_map": color_map,
                                 "generic_keywords": generic_keywords, "product_tax_code": product_tax_code,
                                 "fit_type": fit_type,
                                 "neck_style": neck_style, "pattern_type": pattern_type, "style_name": style_name,
                                 "theme": theme,
                                 "top_style": top_style, "sleeve_type": sleeve_type, "strap_type": strap_type,
                                 "closure_type": closure_type, "outer_material_type3": outer_material_type3,
                                 "material_composition3":
                                     material_composition3, "other_image_url3": other_image_url3,
                                 "other_image_url4": other_image_url4},
                                ignore_index=True)

                            dfObj_2.to_excel(writer, startrow=len(dfObj) + len(dfObj_1) + 2, index=False)
                            writer.save()
                            item_sku_value += 1
                        except:
                            pass
                except:
                    pass
        except:
            print("error: sorry, please check the input Url or internet connection")
            pass
    except:
        pass


if __name__ == '__main__':
    main()
