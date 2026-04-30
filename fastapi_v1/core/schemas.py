from pydantic import BaseModel, Field
from typing import List, Optional, Dict

class LineItem(BaseModel):
    page_number: Optional[int] = Field(None, description="Page number")
    material: Optional[str] = Field(None)
    description: str = Field(description="Name or description")
    quantity: Optional[float] = Field(None)
    uom: Optional[str] = Field(None)
    uom_confidence: Optional[str] = Field(None)
    unit_price: Optional[float] = Field(None)
    line_total: Optional[float] = Field(None)
    line_origin: Optional[str] = Field(None)
    line_destination: Optional[str] = Field(None)

class TaxItem(BaseModel):
    tax_name: str
    tax_amount: float

class FeeItem(BaseModel):
    fee_name: str
    fee_amount: float

class InvoiceData(BaseModel):
    vendor_name: str
    vendor_address: Optional[str] = None
    bill_to: Optional[str] = None
    remit_to: Optional[str] = None
    origin: Optional[str] = None
    origin_confidence: Optional[str] = None
    destination: Optional[str] = None
    destination_confidence: Optional[str] = None
    invoice_number: Optional[str] = None
    invoice_number_confidence: Optional[str] = None
    date: Optional[str] = None
    currency: Optional[str] = None
    subtotal: Optional[float] = None
    taxes: List[TaxItem] = Field(default_factory=list)
    additional_fees: List[FeeItem] = Field(default_factory=list)
    shipping_name: Optional[str] = None
    shipping_handling: Optional[float] = 0.0
    total_amount: float
    total_amount_confidence: Optional[str] = None
    custom_fields: Dict[str, Optional[str]] = Field(default_factory=dict)
    line_items: List[LineItem]

class InvoiceDocument(BaseModel):
    invoices: List[InvoiceData]
