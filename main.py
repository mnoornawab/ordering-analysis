import pandas as pd

# === File Path ===
file_path = 'sample_file.xlsx'  # Replace this with your actual file name if different

# === Step 1: Load Sheets ===
print("📥 Loading Excel file...")
try:
    sima_df = pd.read_excel(file_path, sheet_name='Quantity on Order - SIMA System', engine='openpyxl')
    allocation_df = pd.read_excel(file_path, sheet_name='Allocation File', engine='openpyxl')
    orders_df = pd.read_excel(file_path, sheet_name='Orders-SIMA System', engine='openpyxl')
except Exception as e:
    print("❌ Error loading sheets:", e)
    exit()

print("✅ Loaded sheets successfully.")
print(f"  → SIMA: {sima_df.shape}")
print(f"  → Allocation: {allocation_df.shape}")
print(f"  → Orders: {orders_df.shape}")

# === Step 2: Clean and Extract Columns ===

# --- Clean: Quantity on Order - SIMA System ---
try:
    sima_clean = sima_df[['Item Code', 'Qty On Order']].copy()
    sima_clean.columns = ['item_code', 'qty_on_order']
    sima_clean['item_code'] = sima_clean['item_code'].astype(str).str.strip()
    sima_clean['qty_on_order'] = pd.to_numeric(sima_clean['qty_on_order'], errors='coerce').fillna(0)
except Exception as e:
    print("❌ Error processing SIMA sheet:", e)
    exit()

# --- Clean: Allocation File ---
try:
    allocation_clean = allocation_df[['Material code', 'Pending order qty', 'PO Reference']].copy()
    allocation_clean.columns = ['item_code', 'pending_qty', 'po_reference']
    allocation_clean['item_code'] = allocation_clean['item_code'].astype(str).str.strip()
    allocation_clean['pending_qty'] = pd.to_numeric(allocation_clean['pending_qty'], errors='coerce').fillna(0)
    allocation_clean['po_reference'] = allocation_clean['po_reference'].astype(str).str.strip().str.lower()
except Exception as e:
    print("❌ Error processing Allocation sheet:", e)
    exit()

# --- Clean: Orders-SIMA System ---
try:
    orders_clean = orders_df[['Item Code', 'Total Qty Ordered', 'Reserved', 'Confirmed', 'Balance']].copy()
    orders_clean.columns = ['item_code', 'total_ordered', 'reserved', 'confirmed', 'balance']
    orders_clean['item_code'] = orders_clean['item_code'].astype(str).str.strip()
    for col in ['total_ordered', 'reserved', 'confirmed', 'balance']:
        orders_clean[col] = pd.to_numeric(orders_clean[col], errors='coerce').fillna(0)
except Exception as e:
    print("❌ Error processing Orders sheet:", e)
    exit()

# === Final Confirmation ===
print("\n✅ All sheets cleaned and ready for processing:")
print(f"  → SIMA Cleaned: {sima_clean.shape}")
print(f"  → Allocation Cleaned: {allocation_clean.shape}")
print(f"  → Orders Cleaned: {orders_clean.shape}")
