from woocommerce import API


wcapi = API(
    url="https://ninjagameden.co.uk",
    consumer_key="",
    consumer_secret="",
    query_string_auth=True)

data = {
    "name": "Wipeout Fusion",
    "type": "simple",
    "sku": "BAT115",
    "status": "private",
    "featured": "False",
    "description": "This copy of Wipeout Fusion is in very good clean condition with some minor shelf wear to the "
                   "outer case.",
    "regular_price": "10.00",
    "stock_quantity": "1",
    "categories": [],
    "images": []
}

#print(wcapi.post("products", data).json())

#print(wcapi.delete("products/547", params={"force": True}).json())

data1 = {
    "stock_status": "outofstock"
}

print(wcapi.put("products/547", data1).json())
