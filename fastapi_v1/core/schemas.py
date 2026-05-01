from pydantic import BaseModel, Field
from typing import List, Optional, Dict

class LineItem(BaseModel):
    page_number: Optional[int] = Field(None, description="Page number (starting at 1)")
    material: Optional[str] = Field(None, description="Material code, part number, or SKU")
    description: str = Field(description="Name or description of the item. If a tax or fee is printed as a row INSIDE the main table, extract it here.")
    quantity: Optional[float] = Field(None, description="Number of items SHIPPED. If you see 'QTY B/O' alongside 'QTY SHP', ONLY extract the Shipped amount.")
    uom: Optional[str] = Field(None, description="Unit of Measure (e.g., EA, LBS, KG). Look for headers like 'UOM' or 'Bin'.")
    uom_confidence: Optional[str] = Field(None, description="'High', 'Medium', or 'Low'")
    unit_price: Optional[float] = Field(None, description="Price of a single unit")
    line_total: Optional[float] = Field(None, description="Total cost for this specific line item. If blank, leave null or 0.0.")
    line_origin: Optional[str] = Field(None, description="Exact Origin/Ship-From address printed for THIS specific row.")
    line_destination: Optional[str] = Field(None, description="Exact Destination/Ship-To address printed for THIS specific row.")

class TaxItem(BaseModel):
    tax_name: str = Field(description="The exact printed name of the tax.")
    tax_amount: float = Field(description="The amount for this specific tax.")

class FeeItem(BaseModel):
    fee_name: str = Field(description="The exact printed name of the fee.")
    fee_amount: float = Field(description="The amount for this specific fee.")

class InvoiceData(BaseModel):
    vendor_name: str = Field(description="Name of the company issuing the invoice")
    vendor_address: Optional[str] = Field(None, description="The FULL complete address of the vendor.")
    bill_to: Optional[str] = Field(None, description="The FULL complete 'Bill To' or 'Sold To' address.")
    remit_to: Optional[str] = Field(None, description="The FULL complete 'Remit To' address.")
    origin: Optional[str] = Field(None, description="The FULL origin physical address for the overall invoice.")
    origin_confidence: Optional[str] = Field(None, description="'High', 'Medium', or 'Low'")
    destination: Optional[str] = Field(None, description="The FULL destination physical address for the overall invoice.")
    destination_confidence: Optional[str] = Field(None, description="'High', 'Medium', or 'Low'")
    invoice_number: Optional[str] = Field(None, description="Unique invoice number.")
    invoice_number_confidence: Optional[str] = Field(None, description="'High', 'Medium', or 'Low'")
    date: Optional[str] = Field(None, description="Date the invoice was issued")
    currency: Optional[str] = Field(None, description="3-letter currency code")
    subtotal: Optional[float] = Field(None, description="The subtotal amount before taxes and shipping.")
    taxes: List[TaxItem] = Field(default_factory=list, description="Extract individual taxes ONLY from the summary block at the bottom.")
    additional_fees: List[FeeItem] = Field(default_factory=list, description="ONLY extract fees from the summary block at the bottom.")
    shipping_name: Optional[str] = Field(None, description="The exact printed name of the shipping charge.")
    shipping_handling: Optional[float] = Field(0.0, description="ONLY extract this if it appears in the final summary block at the bottom.")
    total_amount: float = Field(description="Final total amount charged on the invoice")
    total_amount_confidence: Optional[str] = Field(None, description="'High', 'Medium', or 'Low'")
    custom_fields: Dict[str, Optional[str]] = Field(default_factory=dict, description="Extract any custom fields requested by the user.")
    line_items: List[LineItem] = Field(description="List of all individual items purchased.")

class InvoiceDocument(BaseModel):
    invoices: List[InvoiceData] = Field(description="List of distinct invoices.")
