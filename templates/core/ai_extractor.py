import json
import re
from groq import Groq

INVOICE_PROMPT = """You are an expert invoice data extraction AI. Analyze this invoice image and extract ALL key attributes.

Return ONLY a valid JSON object with these fields (use null if not found):
{
  "invoice_number": "", "invoice_date": "", "due_date": "",
  "vendor": {"name": "", "address": "", "email": "", "phone": "", "tax_id": ""},
  "bill_to": {"name": "", "address": "", "email": "", "phone": ""},
  "line_items": [{"description": "", "quantity": "", "unit_price": "", "total": ""}],
  "subtotal": "", "tax_rate": "", "tax_amount": "", "discount": "", "shipping": "",
  "total_amount": "", "currency": "", "payment_terms": "", "payment_method": "",
  "bank_details": "", "notes": "", "po_number": ""
}
Return ONLY the JSON, no markdown, no explanation."""

def extract_invoice_data(api_key, image_b64, mime_type="image/png"):
    client = Groq(api_key=api_key)
    response = client.chat.completions.create(
        model="meta-llama/llama-4-scout-17b-16e-instruct",
        messages=[{
            "role": "user",
            "content": [
                {"type": "image_url", "image_url": {"url": f"data:{mime_type};base64,{image_b64}"}},
                {"type": "text", "text": INVOICE_PROMPT}
            ]
        }],
        max_tokens=2000,
        temperature=0.1
    )
    
    text = response.choices[0].message.content.strip()
    
    # Using hex code \x60 for backticks to prevent markdown cutting off
    text = re.sub(r'^\x60\x60\x60(?:json)?\s*', '', text)
    text = re.sub(r'\s*\x60\x60\x60$', '', text)
    
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        match = re.search(r'\{.*\}', text, re.DOTALL)
        if match:
            return json.loads(match.group())
        raise ValueError("Could not parse AI response into JSON.")