from flask import Flask, render_template_string, request, send_file
import pandas as pd
import io

app = Flask(__name__)

# Pre-defined Client Preset Configurations
CLIENT_PRESETS = {
    "vital": {
        "client_username": "aramex@lavishlivings.com",
        "client_password": "Lavishlivings2026!",
        "account_number": "IDR10632",
        "account_pin": "993744",
        "shipper_name": "Yashasvi Rathore",
        "shipper_company": "VITALKINGDOM",
        "shipper_address1": "Plot no-19, Shiv bhumi-3",
        "shipper_address2": "Nr. Shyam Estate, Bakrol-Kujad Road",
        "shipper_address3": "Bakrol",
        "shipper_city": "Ahmedabad",
        "shipper_state": "Gujrat",
        "shipper_postcode": "382433",
        "shipper_country": "IN",
        "shipper_phone": "6139156710",
        "shipper_email": "operations@lavishlivings.com",
        "hs_code": "21069099",
        "goods_desc": "Suppliments",
        "currency": "USD",
        "customs_value": 99,
        "tax_id": "450714263241"
    },
    "satrajan": {
        "client_username": "aramex@lavishlivings.com",
        "client_password": "Lavishlivings2026!",
        "account_number": "IDR10632",
        "account_pin": "993744",
        "shipper_name": "Ajay Solanki",
        "shipper_company": "Ajay Solanki",
        "shipper_address1": "278-A",
        "shipper_address2": "GULAB BAGH COLONY",
        "shipper_address3": "278-A, GULAB BAGH COLONY",
        "shipper_city": "Indore",
        "shipper_state": "Madhya Pradesh",
        "shipper_postcode": "452010",
        "shipper_country": "IN",
        "shipper_phone": "9685198556",
        "shipper_email": "operations@lavishlivings.com",
        "hs_code": "63041910",
        "goods_desc": "Duvet cover set",
        "currency": "USD",
        "customs_value": 5,
        "tax_id": "594130563238"
    }
}

HTML_FORM = """
<!DOCTYPE html>
<html>
<head>
    <title>Aramex Bulk Label Generator</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
</head>
<body class="bg-light p-4">
<div class="container bg-white p-4 rounded shadow-sm" style="max-width: 700px;">
    <h3 class="mb-3 text-primary">📦 Aramex Bulk Label Generator</h3>
    <form method="POST" action="/generate">
        <div class="mb-3">
            <label class="form-label fw-bold">Select Client Profile</label>
            <select name="client_type" class="form-select" required>
                <option value="vital">Vital Kingdom (Suppliments)</option>
                <option value="satrajan">Satrajan (Duvet cover set)</option>
            </select>
        </div>
        <hr>
        <h5 class="text-secondary">Consignee & Package Details</h5>
        <div class="row">
            <div class="col-md-6 mb-3">
                <label class="form-label">Invoice Number Prefix/Start</label>
                <input type="text" name="invoice_no" class="form-control" placeholder="VK-2627-IN-023" required>
            </div>
            <div class="col-md-6 mb-3">
                <label class="form-label">Invoice Date (MM/DD/YYYY)</label>
                <input type="text" name="invoice_date" class="form-control" value="07/29/2026" required>
            </div>
        </div>
        <div class="row">
            <div class="col-md-6 mb-3">
                <label class="form-label">Receiver Name</label>
                <input type="text" name="receiver_name" class="form-control" placeholder="Abdul Rahman" required>
            </div>
            <div class="col-md-6 mb-3">
                <label class="form-label">Receiver Phone</label>
                <input type="text" name="receiver_phone" class="form-control" placeholder="564189901" required>
            </div>
        </div>
        <div class="mb-3">
            <label class="form-label">Receiver Address Line 1</label>
            <input type="text" name="receiver_addr1" class="form-control" placeholder="54 Amman St, Al Qusais Ind. Third" required>
        </div>
        <div class="row">
            <div class="col-md-6 mb-3">
                <label class="form-label">Receiver City</label>
                <input type="text" name="receiver_city" class="form-control" value="Dubai" required>
            </div>
            <div class="col-md-6 mb-3">
                <label class="form-label">Country Code</label>
                <input type="text" name="receiver_country" class="form-control" value="AE" required>
            </div>
        </div>
        <hr>
        <div class="row">
            <div class="col-md-4 mb-3">
                <label class="form-label fw-bold">Total Shipments (Rows)</label>
                <input type="number" name="total_shipments" class="form-control" value="1" min="1" required>
            </div>
            <div class="col-md-4 mb-3">
                <label class="form-label fw-bold">Boxes per AWB (DO Col)</label>
                <input type="number" name="num_pieces" class="form-control" value="3" min="1" required>
            </div>
            <div class="col-md-4 mb-3">
                <label class="form-label fw-bold">Weight (KG)</label>
                <input type="number" step="0.01" name="weight" class="form-control" value="37.56" required>
            </div>
        </div>
        <button type="submit" class="btn btn-primary w-100 fw-bold mt-3">🚀 Download Aramex Excel File</button>
    </form>
</div>
</body>
</html>
"""

@app.route("/")
def index():
    return render_template_string(HTML_FORM)

@app.route("/generate", methods=["POST"])
def generate():
    client_key = request.form.get("client_type")
    preset = CLIENT_PRESETS[client_key]
    
    total_shipments = int(request.form.get("total_shipments", 1))
    num_pieces = int(request.form.get("num_pieces", 1))
    weight = float(request.form.get("weight", 1.0))
    
    inv_start = request.form.get("invoice_no")
    inv_date = request.form.get("invoice_date")
    rec_name = request.form.get("receiver_name")
    rec_phone = request.form.get("receiver_phone")
    rec_addr1 = request.form.get("receiver_addr1")
    rec_city = request.form.get("receiver_city")
    rec_country = request.form.get("receiver_country")
    
    # 162 Columns Structure
    cols = [
        'ClientInfo: UserName', 'ClientInfo: Password', 'ClientInfo: Version', 'ClientInfo: AccountNumber', 'ClientInfo: AccountPin', 'ClientInfo: AccountEntity', 'ClientInfo: AccountCountryCode', 'ClientInfo: Source', 'LabelInfo: ReportID', 'LabelInfo: ReportType', 'Shipments: Reference1', 'Shipments: Reference2', 'Shipments: Reference3', 'Shipments: Shipper: Reference1', 'Shipments: Shipper: Reference2', 'Shipments: Shipper: AccountNumber', 'Shipments: Shipper: PartyAddress: Line1', 'Shipments: Shipper: PartyAddress: Line2', 'Shipments: Shipper: PartyAddress: Line3', 'Shipments: Shipper: PartyAddress: City', 'Shipments: Shipper: PartyAddress: StateOrProvinceCode', 'Shipments: Shipper: PartyAddress: PostCode', 'Shipments: Shipper: PartyAddress: CountryCode', 'Shipments: Shipper: PartyAddress: Longitude', 'Shipments: Shipper: PartyAddress: Latitude', 'Shipments: Shipper: PartyAddress: BuildingNumber', 'Shipments: Shipper: PartyAddress: BuildingName', 'Shipments: Shipper: PartyAddress: Floor', 'Shipments: Shipper: PartyAddress: Apartment', 'Shipments: Shipper: PartyAddress: POBox', 'Shipments: Shipper: PartyAddress: Description', 'Shipments: Shipper: Contact: Department', 'Shipments: Shipper: Contact: PersonName', 'Shipments: Shipper: Contact: Title', 'Shipments: Shipper: Contact: CompanyName', 'Shipments: Shipper: Contact: PhoneNumber1', 'Shipments: Shipper: Contact: PhoneNumber1Ext', 'Shipments: Shipper: Contact: PhoneNumber2', 'Shipments: Shipper: Contact: PhoneNumber2Ext', 'Shipments: Shipper: Contact: FaxNumber', 'Shipments: Shipper: Contact: CellPhone', 'Shipments: Shipper: Contact: EmailAddress', 'Shipments: Shipper: Contact: Type', 'Shipments: Consignee: Reference1', 'Shipments: Consignee: Reference2', 'Shipments: Consignee: AccountNumber', 'Shipments: Consignee: PartyAddress: Line1', 'Shipments: Consignee: PartyAddress: Line2', 'Shipments: Consignee: PartyAddress: Line3', 'Shipments: Consignee: PartyAddress: City', 'Shipments: Consignee: PartyAddress: StateOrProvinceCode', 'Shipments: Consignee: PartyAddress: PostCode', 'Shipments: Consignee: CountryCode', 'Shipments: Consignee: PartyAddress: Longitude', 'Shipments: Consignee: PartyAddress: Latitude', 'Shipments: Consignee: PartyAddress: BuildingNumber', 'Shipments: Consignee: PartyAddress: BuildingName', 'Shipments: Consignee: PartyAddress: Floor', 'Shipments: Consignee: PartyAddress: Apartment', 'Shipments: Consignee: PartyAddress: POBox', 'Shipments: Consignee: PartyAddress: Description', 'Shipments: Consignee: Contact: Department', 'Shipments: Consignee: Contact: PersonName', 'Shipments: Consignee: Contact: Title', 'Shipments: Consignee: Contact: CompanyName', 'Shipments: Consignee: Contact: PhoneNumber1', 'Shipments: Consignee: Contact: PhoneNumber1Ext', 'Shipments: Consignee: Contact: PhoneNumber2', 'Shipments: Consignee: Contact: PhoneNumber2Ext', 'Shipments: Consignee: Contact: FaxNumber', 'Shipments: Consignee: Contact: CellPhone', 'Shipments: Consignee: Contact: EmailAddress', 'Shipments: Consignee: Contact: Type', 'Shipments: ThirdParty: Reference1', 'Shipments: ThirdParty: Reference2', 'Shipments: ThirdParty: AccountNumber', 'Shipments: ThirdParty: PartyAddress: Line1', 'Shipments: ThirdParty: PartyAddress: Line2', 'Shipments: ThirdParty: PartyAddress: Line3', 'Shipments: ThirdParty: PartyAddress: City', 'Shipments: ThirdParty: PartyAddress: StateOrProvinceCode', 'Shipments: ThirdParty: PartyAddress: PostCode', 'Shipments: ThirdParty: PartyAddress: CountryCode', 'Shipments: ThirdParty: PartyAddress: Longitude', 'Shipments: ThirdParty: PartyAddress: Latitude', 'Shipments: ThirdParty: PartyAddress: BuildingNumber', 'Shipments: ThirdParty: PartyAddress: BuildingName', 'Shipments: ThirdParty: PartyAddress: Floor', 'Shipments: ThirdParty: PartyAddress: Apartment', 'Shipments: ThirdParty: PartyAddress: POBox', 'Shipments: ThirdParty: PartyAddress: Description', 'Shipments: ThirdParty: Contact: Department', 'Shipments: ThirdParty: Contact: PersonName', 'Shipments: ThirdParty: Contact: Title', 'Shipments: ThirdParty: Contact: CompanyName', 'Shipments: ThirdParty: Contact: PhoneNumber1', 'Shipments: ThirdParty: Contact: PhoneNumber1Ext', 'Shipments: ThirdParty: Contact: PhoneNumber2', 'Shipments: ThirdParty: Contact: PhoneNumber2Ext', 'Shipments: ThirdParty: Contact: FaxNumber', 'Shipments: ThirdParty: Contact: CellPhone', 'Shipments: ThirdParty: Contact: EmailAddress', 'Shipments: ThirdParty: Contact: Type', 'Shipments: ShippingDateTime', 'Shipments: DueDate', 'Shipments: Comments', 'Shipments: PickupLocation', 'Shipments: OperationsInstructions', 'Shipments: AccountingInstrcutions', 'Shipments: Details: Dimensions: Length', 'Shipments: Details: Dimensions: Width', 'Shipments: Details: Dimensions: Height', 'Shipments: Details: Dimensions: Unit', 'Shipments: Details: ActualWeight: Unit', 'Shipments: Details: ActualWeight: Value', 'Shipments: Details: ChargeableWeight', 'Shipments: Details: DescriptionOfGoods', 'Shipments: Details: GoodsOriginCountry', 'Shipments: Details: NumberOfPieces', 'Shipments: Details: ProductGroup', 'Shipments: Details: ProductType', 'Shipments: Details: PaymentType', 'Shipments: Details: PaymentOptions', 'Shipments: Details: CustomsValueAmount: CurrencyCode', 'Shipments: Details: CustomsValueAmount: Value', 'Shipments: Details: CashAdditionalAmountDescription', 'Shipments: Details: Services', 'Shipments: Details: Items: PackageType', 'Shipments: Details: Items: Quantity', 'Shipments: Details: Items: Weight', 'Shipments: Details: Items: CustomsValue: CurrencyCode', 'Shipments: Details: Items: CustomsValue: Value', 'Shipments: Details: Items: Comments', 'Shipments: Details: Items: GoodsDescription', 'Shipments: Details: Items: CountryOfOrigin', 'Shipments: Details: Items: Reference', 'Shipments: Details: Items: CommodityCode', 'Shipments: Details: AdditionalProperties: CategoryName: CustomsClearance', 'Shipments: Details: AdditionalProperties: CategoryName: CustomsClearance-1', 'Shipments: Details: AdditionalProperties: CategoryName: CustomsClearance-2', 'Shipments: Details: AdditionalProperties: CategoryName: CustomsClearance-3', 'Shipments: Details: AdditionalProperties: CategoryName: CustomsClearance-4', 'Shipments: Details: AdditionalProperties: CategoryName: CustomsClearance-5', 'Shipments: Details: AdditionalProperties: CategoryName: CustomsClearance-6', 'Shipments: Details: AdditionalProperties: Name: ShipperTaxIdVATEINNumber', 'Shipments: Details: AdditionalProperties: Name: ConsigneeTaxIdVATEINNumber', 'Shipments: Details: AdditionalProperties: Name: TaxPaid', 'Shipments: Details: AdditionalProperties: Name: TaxAmount', 'Shipments: Details: AdditionalProperties: Name: InvoiceDate', 'Shipments: Details: AdditionalProperties: Name: InvoiceNumber', 'Shipments: Details: AdditionalProperties: Name: ExporterType', 'Shipments: Attachments', 'Shipments: ForeignHAWB', 'Shipments: TransportType', 'Shipments: PickupGUID', 'Shipments: Number', 'Shipments: ScheduledDelivery', 'Transaction: Reference1', 'Transaction: Reference2', 'Transaction: Reference3', 'Transaction: Reference4', 'Transaction: Reference5'
    ]
    
    rows = []
    for i in range(total_shipments):
        inv_no = f"{inv_start}-{i+1:02d}" if total_shipments > 1 else inv_start
        
        row_dict = {col: None for col in cols}
        
        # Client Info
        row_dict['ClientInfo: UserName'] = preset['client_username']
        row_dict['ClientInfo: Password'] = preset['client_password']
        row_dict['ClientInfo: Version'] = 'v1.0'
        row_dict['ClientInfo: AccountNumber'] = preset['account_number']
        row_dict['ClientInfo: AccountPin'] = preset['account_pin']
        row_dict['ClientInfo: AccountEntity'] = 'IDR'
        row_dict['ClientInfo: AccountCountryCode'] = 'IN'
        row_dict['ClientInfo: Source'] = 24
        row_dict['LabelInfo: ReportID'] = 9729
        row_dict['LabelInfo: ReportType'] = 'URL'
        
        # Shipper Info
        row_dict['Shipments: Shipper: AccountNumber'] = preset['account_number']
        row_dict['Shipments: Shipper: PartyAddress: Line1'] = preset['shipper_address1']
        row_dict['Shipments: Shipper: PartyAddress: Line2'] = preset['shipper_address2']
        row_dict['Shipments: Shipper: PartyAddress: Line3'] = preset['shipper_address3']
        row_dict['Shipments: Shipper: PartyAddress: City'] = preset['shipper_city']
        row_dict['Shipments: Shipper: PartyAddress: StateOrProvinceCode'] = preset['shipper_state']
        row_dict['Shipments: Shipper: PartyAddress: PostCode'] = preset['shipper_postcode']
        row_dict['Shipments: Shipper: PartyAddress: CountryCode'] = preset['shipper_country']
        row_dict['Shipments: Shipper: PartyAddress: Longitude'] = 0
        row_dict['Shipments: Shipper: PartyAddress: Latitude'] = 0
        row_dict['Shipments: Shipper: Contact: PersonName'] = preset['shipper_name']
        row_dict['Shipments: Shipper: Contact: CompanyName'] = preset['shipper_company']
        row_dict['Shipments: Shipper: Contact: PhoneNumber1'] = preset['shipper_phone']
        row_dict['Shipments: Shipper: Contact: CellPhone'] = preset['shipper_phone']
        row_dict['Shipments: Shipper: Contact: EmailAddress'] = preset['shipper_email']
        
        # Consignee Info
        row_dict['Shipments: Consignee: Reference1'] = inv_no
        row_dict['Shipments: Consignee: PartyAddress: Line1'] = rec_addr1
        row_dict['Shipments: Consignee: PartyAddress: City'] = rec_city
        row_dict['Shipments: Consignee: CountryCode'] = rec_country
        row_dict['Shipments: Consignee: PartyAddress: Longitude'] = 0
        row_dict['Shipments: Consignee: PartyAddress: Latitude'] = 0
        row_dict['Shipments: Consignee: Contact: PersonName'] = rec_name
        row_dict['Shipments: Consignee: Contact: PhoneNumber1'] = rec_phone
        row_dict['Shipments: Consignee: Contact: CellPhone'] = rec_phone
        row_dict['Shipments: Consignee: Contact: EmailAddress'] = 'operations@lavishlivings.com'
        
        # Shipment Details
        row_dict['Shipments: Details: Dimensions: Length'] = 58 if client_key == 'vital' else 1
        row_dict['Shipments: Details: Dimensions: Width'] = 38 if client_key == 'vital' else 1
        row_dict['Shipments: Details: Dimensions: Height'] = 32 if client_key == 'vital' else 1
        row_dict['Shipments: Details: Dimensions: Unit'] = 'CM'
        row_dict['Shipments: Details: ActualWeight: Unit'] = 'KG'
        row_dict['Shipments: Details: ActualWeight: Value'] = weight
        row_dict['Shipments: Details: DescriptionOfGoods'] = preset['goods_desc']
        row_dict['Shipments: Details: GoodsOriginCountry'] = 'IN'
        
        # Column DO (NumberOfPieces) Logic
        row_dict['Shipments: Details: NumberOfPieces'] = num_pieces
        
        row_dict['Shipments: Details: ProductGroup'] = 'EXP'
        row_dict['Shipments: Details: ProductType'] = 'PPX'
        row_dict['Shipments: Details: PaymentType'] = 'P'
        row_dict['Shipments: Details: CustomsValueAmount: CurrencyCode'] = preset['currency']
        row_dict['Shipments: Details: CustomsValueAmount: Value'] = preset['customs_value']
        row_dict['Shipments: Details: Services'] = 'FRDM'
        row_dict['Shipments: Details: Items: Quantity'] = 1
        row_dict['Shipments: Details: Items: CustomsValue: CurrencyCode'] = preset['currency']
        row_dict['Shipments: Details: Items: CustomsValue: Value'] = preset['customs_value']
        row_dict['Shipments: Details: Items: GoodsDescription'] = preset['goods_desc']
        row_dict['Shipments: Details: Items: CountryOfOrigin'] = 'IN'
        row_dict['Shipments: Details: Items: CommodityCode'] = preset['hs_code']
        
        row_dict['Shipments: Details: AdditionalProperties: Name: ShipperTaxIdVATEINNumber'] = preset['tax_id']
        row_dict['Shipments: Details: AdditionalProperties: Name: TaxPaid'] = 'NO'
        row_dict['Shipments: Details: AdditionalProperties: Name: InvoiceDate'] = inv_date
        row_dict['Shipments: Details: AdditionalProperties: Name: InvoiceNumber'] = inv_no
        row_dict['Shipments: Details: AdditionalProperties: Name: ExporterType'] = 'UT'
        row_dict['Shipments: TransportType'] = 0
        
        rows.append(row_dict)
        
    df = pd.DataFrame(rows)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    output.seek(0)
    
    return send_file(output, download_name=f"{inv_start}_aramex_bulk.xlsx", as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
