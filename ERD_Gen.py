import requests
import json
from PIL import Image

def mermaid_to_image(mermaid_code: str, output_filename: str = "diagram.png"):
    """
    Converts a Mermaid.js diagram string into an image with a white background using the Kroki API.
    
    Args:
        mermaid_code (str): The Mermaid.js diagram definition.
        output_filename (str): The filename to save the generated image (default: "diagram.png").
    
    Returns:
        str: The file path of the saved image if successful, otherwise an error message.
    """
    # Send request to Kroki API
    url = "https://kroki.io/mermaid/png"
    headers = {"Content-Type": "application/json"}
    data = json.dumps({"diagram_source": mermaid_code})

    response = requests.post(url, headers=headers, data=data)

    # Save and modify the image if the request is successful
    if response.status_code == 200:
        temp_filename = "temp_diagram.png"
        with open(temp_filename, "wb") as f:
            f.write(response.content)
        
        # Open image and convert to white background
        img = Image.open(temp_filename).convert("RGBA")
        white_bg = Image.new("RGB", img.size, (255, 255, 255))  # White background
        white_bg.paste(img, mask=img.split()[3])  # Apply alpha mask
        
        # Save final image
        white_bg.save(output_filename, "PNG")
        return f"✅ ER Diagram saved with white background as '{output_filename}'"
    else:
        return f"❌ Failed to generate diagram. API Response: {response.text}"

# # Example usage
# mermaid_code = """
# erDiagram
#     CUSTOMERS_ {
#         NUMBER CUSTOMER_ID PK
#         NUMBER AGE
#         VARCHAR GENDER
#         VARCHAR STATE
#         VARCHAR CITY
#         VARCHAR POSTAL_CODE
#         NUMBER LATITUDE
#         NUMBER LONGITUDE
#     }
#     GEO_ {
#         VARCHAR STATE
#         VARCHAR ABBREVIATION PK
#         VARCHAR REGION
#         NUMBER LATITUDE
#         NUMBER LONGITUDE
#     }
#     PRODUCT_ {
#         NUMBER PRODUCT_ID PK
#         VARCHAR NAME
#         VARCHAR CATEGORY
#         VARCHAR CATEGORY_HEAD
#         VARCHAR BRAND
#         VARCHAR DEPARTMENT
#         NUMBER SHIPPING_COST_1000_MILE
#         NUMBER RETAIL_PRICE
#     }
#     RBAC_CUSTOMERS_ {
#         NUMBER CUSTOMER_ID PK
#         NUMBER AGE
#         VARCHAR GENDER
#         VARCHAR STATE
#         VARCHAR CITY
#         VARCHAR POSTAL_CODE
#         NUMBER LATITUDE
#         NUMBER LONGITUDE
#     }
#     RBAC_PRODUCT_ {
#         NUMBER PRODUCT_ID PK
#         VARCHAR NAME
#         VARCHAR CATEGORY
#         VARCHAR CATEGORY_HEAD
#         VARCHAR BRAND
#         VARCHAR DEPARTMENT
#         NUMBER SHIPPING_COST_1000_MILE
#         NUMBER RETAIL_PRICE
#     }
#     RBAC_SALES_ {
#         NUMBER ORDER_ITEM_ID PK
#         NUMBER PRODUCT_ID FK
#         TIMESTAMP_TZ TRANSACTION_DATE
#         NUMBER QUANTITY
#         NUMBER SALES
#         NUMBER CUSTOMER_ID FK
#         VARCHAR STATE_AB FK
#     }
#     SALES_ {
#         NUMBER ORDER_ITEM_ID PK
#         NUMBER PRODUCT_ID
#         TIMESTAMP_TZ TRANSACTION_DATE
#         NUMBER QUANTITY
#         NUMBER SALES
#         NUMBER CUSTOMER_ID
#         VARCHAR STATE_AB
#     }

#     RBAC_SALES_ ||--o| RBAC_PRODUCT_ : "PRODUCT_ID (Many-to-One)"
#     RBAC_SALES_ ||--o| RBAC_CUSTOMERS_ : "CUSTOMER_ID (Many-to-One)"
#     RBAC_SALES_ ||--o| GEO_ : "STATE_AB (Many-to-One)"
#     SALES_ ||--o| PRODUCT_ : "PRODUCT_ID (Many-to-One)"
#     SALES_ ||--o| CUSTOMERS_ : "CUSTOMER_ID (Many-to-One)"
#     SALES_ ||--o| GEO_ : "STATE_AB (Many-to-One)"
# """

# # Call the function
# print(mermaid_to_image(mermaid_code, "er_diagram_white.png"))