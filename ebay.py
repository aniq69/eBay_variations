import json
import requests
from parsel import Selector
from collections import defaultdict
import openpyxl
from tqdm import tqdm


def find_json_objects(text: str, decoder=json.JSONDecoder()):
  pos = 0
  while True:
    match = text.find("{", pos)
    if match == -1:
      break
    try:
      result, index = decoder.raw_decode(text[match:])
      yield result
      pos = match + index
    except ValueError:
      pos = match + 1


def parse_variants(sel: Selector) -> dict:
  script = sel.xpath('//script[contains(., "itemVariationsMap")]/text()').get()
  if not script:
    return {}

  all_data = list(find_json_objects(script))
  variants = next((d for d in all_data if "itemVariationsMap" in str(d)),
                  {}).get("itemVariationsMap", {})

  selections = defaultdict(dict)
  for selection in sel.css(".x-msku__box-cont select"):
    name = selection.xpath("@selectboxlabel").get()
    selection_data = {}
    for option in selection.xpath("option"):
      value = int(option.xpath("@value").get())
      if value == -1:
        continue
      label = option.xpath("text()").get().strip()
      label = label.split("(Out ")[0]
      selections[name][value] = label

  for variant_id, variant in variants.items():
    for trait, trait_id in variant["traitValuesMap"].items():
      variant["traitValuesMap"][trait] = selections[trait][trait_id]

  parsed_variants = {}
  for variant_id, variant in variants.items():
    label = " ".join(variant["traitValuesMap"].values())
    parsed_variants[label] = {
        "id": variant_id,
        "price": variant["price"],
        "price_converted": variant["convertedPrice"],
        "vat_price": variant["vatPrice"],
        "quantity": variant["quantity"],
        "in_stock": variant["inStock"],
        "sold": variant["quantitySold"],
        "available": variant["quantityAvailable"],
        "watch_count": variant["watchCount"],
        "epid": variant["epid"],
        "top_product": variant["topProduct"],
        "traits": variant["traitValuesMap"],
    }
  return parsed_variants


def read_product_links(filename):
  with open(filename, 'r') as file:
    return file.readlines()


def update_file(filename, links):
  with open(filename, 'w') as file:
    for link in links:
      file.write(link + '\n')


def main():
  input_filename = 'links.txt'
  output_filename = 'output.xlsx'
  variation_not_found_filename = 'variation_not_found.txt'
  variation_found_filename = 'variation_found.txt'

  product_links = read_product_links(input_filename)
  links_with_variations = []
  links_without_variations = []

  workbook = openpyxl.Workbook()
  worksheet = workbook.active
  fieldnames = [
      "Label", "id", "price", "price_converted", "vat_price", "quantity",
      "in_stock", "sold", "available", "watch_count", "epid", "top_product",
      "traits", "eBay Item Number"
  ]
  worksheet.append(fieldnames)
  workbook.save(output_filename)

  for index, link in enumerate(
      tqdm(product_links, desc="Scraping", unit="product")):
    response = requests.get(link.strip())

    if response.status_code != 200:
      print(f"Failed to fetch data for product {index + 1}")
      continue

    sel = Selector(response.text)
    variants_data = parse_variants(sel)

    if variants_data:
      links_with_variations.append(link.strip())
    else:
      links_without_variations.append(link.strip())

    for label, variant_data in variants_data.items():
      ebay_item_number = link.strip().split('/')[-1].split('?')[0]
      row_data = [
          label, variant_data["id"], variant_data["price"],
          variant_data["price_converted"], variant_data["vat_price"],
          variant_data["quantity"], variant_data["in_stock"],
          variant_data["sold"], variant_data["available"],
          variant_data["watch_count"], variant_data["epid"],
          variant_data["top_product"],
          json.dumps(variant_data["traits"]), ebay_item_number
      ]
      worksheet.append(row_data)
      workbook.save(output_filename)

    update_file(variation_found_filename, links_with_variations)
    update_file(variation_not_found_filename, links_without_variations)

  print(f"Links this much: {len(product_links)}")
  print(f"Found variation in these links: {len(links_with_variations)}")
  print(
      f"And these links ({len(links_without_variations)}) are without variations"
  )


if __name__ == "__main__":
  main()