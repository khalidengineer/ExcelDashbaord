# 🚀 TIOMP — 2 Minute Setup (Hinglish)

Bas **ek single VBA file** paste karni hai. Multiple files import karne ki zarurat nahi.

---

## ✅ Aapko Sirf Ye Karna Hai

### Step 1 — File Download Karo
GitHub se ye **EK FILE** download karo:

**[`vba/TIOMP_AllInOne.bas`](vba/TIOMP_AllInOne.bas)** ← bas yahi ek file

(Right-click → Save Link As, ya GitHub par file kholke `Download` button)

### Step 2 — Excel Workbook Banao
1. Excel kholo, blank workbook banao
2. **`File → Save As`** → naam de **`TIOMP.xlsm`**
3. **Important:** `Apps.csv` jis folder me hai, **ussi folder me save karo**

```
Aapka folder kuch aisa dikhna chahiye:
📁 MyDashboard\
   ├── Apps.csv
   └── TIOMP.xlsm
```

### Step 3 — Macros Enable Karo (pehli baar)
1. **`File → Options → Trust Center → Trust Center Settings`**
2. **`Macro Settings`** → "Enable VBA macros" select karo
3. Niche checkbox **"Trust access to the VBA project object model"** ✅
4. **OK → OK**

### Step 4 — Code Paste Karo
1. Excel me **`Alt + F11`** dabao (VBA editor khulega)
2. Top menu se **`Insert → Module`**
3. Right side me khali code window khulegi
4. `TIOMP_AllInOne.bas` file ko Notepad ya VS Code me kholo
5. **`Ctrl + A`** → **`Ctrl + C`** (poori file copy)
6. VBA window me click karke **`Ctrl + V`** (paste)
7. **`Ctrl + S`** save karo

### Step 5 — Run Karo!
1. **`Alt + F8`** dabao
2. List me **`BuildEnterpriseDashboard`** select karo
3. **`Run`** click karo

⏱️ **2-5 second me dashboard ready!** 🎉

---

## 🎯 Kya Milega

Auto-build ho jayega:
- **16 sheets** (Dashboard + 8 modules + Settings + Logs + 5 hidden)
- **8 hero KPI cards** + 10 secondary KPIs
- **13+ dark-themed charts** (trend, donut, gauge, radar, heatmap)
- **AI Alert Feed** with colored alerts
- **Live Activity Console**
- **Side navigation** for all modules
- **5 buttons**: REFRESH, EXPORT PDF, EMAIL MIS, THEME, SEARCH

---

## 🔄 Roz Kaise Update Karenge

`Apps.csv` me nayi rows add ho gayi? Bas dashboard me **REFRESH** button click karo. Bas.

---

## ❌ Problem? Common Fixes

| Problem | Solution |
|---------|----------|
| `Compile error` | Pura code copy karna bhul gaye, dobara `Ctrl+A` se sab select karke paste karo |
| `Apps.csv not found` | `Apps.csv` ko `TIOMP.xlsm` ke same folder me rakho |
| KPIs me sirf 0 | `RefreshAll` macro chalao |
| Macros disabled warning | Step 3 wapas check karo |
| Code paste karne pe yellow line aaye | Already paste ho gaya hoga, save karke run karo |

---

## 💡 Pro Tip

**Aap purane multi-file approach (M01-M07) bhi use kar sakte ho** agar customize karna hai. Wo files bhi `vba/` folder me hain. Lekin agar bas use karna hai, to **`TIOMP_AllInOne.bas`** ek hi file paste karo, kaam ho gaya.
