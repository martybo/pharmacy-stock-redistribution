# 🏥 Pharmacy Dead Stock Allocator

**Version:** v13.2  
**Language:** Python 3.9+  
**Maintainer:** MediCare Pharmacy Group  
**Primary branch:** `main` (deployment)

---

## 📦 Overview

The **Pharmacy Dead Stock Allocator** automates the redistribution of unused ("dead") stock between community pharmacies within the MediCare Pharmacy Group.  
It ensures medicines identified as *at risk of expiry* are transferred to sites where they are most likely to be used before expiry — reducing waste, freeing stock holding, and improving stock efficiency.

---

## ⚙️ Core Logic

Each stock item identified from the annual stock-take is assessed against dispensing activity across all branches.

The allocator then:

1. **Identifies "dead stock"**  
   – Products that have had **no dispensing in the past 4 months** at the source branch.  
2. **Calculates potential receivers**  
   – Branches that have dispensed that product (or equivalents) within the last 4 months.  
3. **Caps allocation**  
   – Each receiving branch can receive up to its **4-month usage total**,  
     minus a small baseline of **1 pack** if no stock-take data is recorded for that site.  
4. **Allocates by priority:**
   1. **NHD priority:** same-site Nursing Home Dispensary (NHD) if usage and capacity allow.  
   2. **Same Area Manager (AM):** branches under the same AM group.  
   3. **Cross AM:** any remaining stock redistributed across the wider group.  
5. **Balances fairly**  
   – Within each priority group, allocation is *weighted by usage* so higher-need sites get proportionally more.  
6. **Leaves unallocated stock** if no suitable receivers remain — this is logged for manual review.

---

## 🧮 Output Structure

All outputs are generated into a folder:
