# Excel AI Test Report: `R4 Stock Management.xlsm`
**Timestamp:** 2026-10-14T23:17:07.787196
**Sandbox Copy (deleted after test):** `C:\Users\mattc\AppData\Local\Temp\QPMS-Consult-test-runs\R4 Stock Management_test_20261014_231707_844162.xlsm`

**Result:** FAILED

## Test Scenario: `R4 stock adjustment write-off`
**Status:** ❌ FAILED

### Macro Execution
- **ButtonUtils.ButtSaveStockAdjustment**: ❌ Success=False
  - **Error:** `Could not safely dismiss Excel dialog 'Adjustment Complete'.; Expected Excel dialog 'Adjustment Complete' with OK 1 more time(s), but it did not appear.`

### CSV Data Assertions
- `Material Management/Tables/RM_Trans.csv`: ✅ matched 1 row(s); expected 1.
- `Material Management/Tables/QC_Number.csv`: ❌ matched 0 row(s); expected 1.

### ⚠️ Errors Detected (1)
- **Material Management/Tables/QC_Number.csv**: `Expected 1 matching CSV row(s), found 0.` (Formula: `None`)

### Sheet Snapshot: `Stock_Adjust` (R1C2:R24C13)
| Row | Col 1 | Col 2 | Col 3 | Col 4 | Col 5 | Col 6 | Col 7 | Col 8 | Col 9 | Col 10 | Col 11 | Col 12 |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 |  |  | Raw Material Stock Adjustments (Adhoc) |  |  |  |  |  |  |  |  |  |
| 2 |  |  |  |  |  |  |  |  |  |  |  |  |
| 3 |  |  |  |  |  |  |  |  |  |  |  |  |
| 4 |  | Standard Name: |  | * |  |  | QC Number | Retest Date | Nil Stock? | QC Balance | UoM | Status |
| 5 |  | RM Code: |  | * |  |  | 240087 | 31/07/2028 | Yes | 0 | mL | Released |
| 6 |  | QC Number: |  | * |  |  | 250085 | 31/05/2029 | No | 4828.105 | mL | Sampled |
| 7 |  | QC Balance: |  |  |  |  |  |  |  |  |  |  |
| 8 |  | Master UoM: | `=IF(OR(D5="",D6=""),"",INDEX(RM_Master!K:K,MATCH(Stock_Adjust!D5,RM_Master!A:A,0)))` |  |  |  |  |  |  |  |  |  |
| 9 |  | QC Status: | `=IF(D6="","",INDEX(QC_Number!AH:AH,MATCH(D6,QC_Number!A:A,0)))` |  |  |  |  |  |  |  |  |  |
| 10 |  | Nil Stock Flag: | `=IF(D6="","",IF(INDEX(QC_Number!AE:AE,MATCH(D6,QC_Number!A:A,0))=0,"N",INDEX(QC_Number!AE:AE,MATCH(D6,QC_Number!A:A,0))))` |  |  |  |  |  |  |  |  |  |
| 11 |  | Adjustment Type: |  | * |  |  |  |  |  |  |  |  |
| 12 |  | Adjustment Quantity: |  | * |  |  |  |  |  |  |  |  |
| 13 |  | Adjustment UoM: |  | * |  |  |  |  |  |  |  |  |
| 14 |  | New QC Balance: | `=IF(OR(D6="",D12="",D11=""),"",IF(D11="Write-Off",IF((ROUND(D7,3)-ROUND(D12,3))<0,"Insufficient Stock",ROUND(D7,3)-ROUND(D12,3)),ROUND(D7,3)+ROUND(D12,3)))` |  |  |  |  |  |  |  |  |  |
| 15 |  | Adjustment Reason: |  | * |  |  |  |  |  |  |  |  |

### Sheet Snapshot: `Stocktake_QC_Number` (R1C2:R23C13)
| Row | Col 1 | Col 2 | Col 3 | Col 4 | Col 5 | Col 6 | Col 7 | Col 8 | Col 9 | Col 10 | Col 11 | Col 12 |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 |  |  | Stocktake by QC number |  |  |  |  |  |  |  |  |  |
| 2 |  |  |  |  |  |  |  |  |  |  |  |  |
| 3 |  |  |  |  |  |  |  |  |  |  |  |  |
| 4 |  | Standard Name: |  |  |  |  |  |  |  |  |  |  |
| 5 |  | RM Code: |  |  |  |  |  |  |  |  |  |  |
| 6 |  | QC Number: |  | * |  |  |  |  |  |  |  |  |
| 7 |  | QC Balance: | `=IF(D6="","",INDEX(QC_Number!U:U,MATCH(D6,QC_Number!A:A,0)))` |  |  |  |  |  |  |  |  |  |
| 8 |  | Master UoM: | `=IF(D5="","",INDEX(RM_Master!K:K,MATCH(D5,RM_Master!A:A,0)))` |  |  |  |  |  |  |  |  |  |
| 9 |  | Count Quantity: |  | * |  |  |  |  |  |  |  |  |
| 10 |  | Count UoM: |  | * |  |  |  |  |  |  |  |  |
| 11 |  | Difference: | `=IF(D9="","",D9-D7)` |  |  |  |  |  |  |  |  |  |
| 12 |  | Expiry/Retest Date: | `=IF(D6="","",INDEX(QC_Number!AC:AC,MATCH(Stocktake_QC_Number!D6,QC_Number!A:A,0)))` |  |  |  |  |  |  |  |  |  |
| 13 |  | Actual Expiry/Retest Date: |  | Only required if different. |  |  |  |  |  |  |  |  |
| 14 |  | Notes: |  |  |  |  |  |  |  |  |  |  |
| 15 |  |  |  |  |  |  |  |  |  |  |  |  |

### Sheet Snapshot: `Stocktake_Sheet` (R1C3:R30C17)
| Row | Col 1 | Col 2 | Col 3 | Col 4 | Col 5 | Col 6 | Col 7 | Col 8 | Col 9 | Col 10 | Col 11 | Col 12 | Col 13 | Col 14 | Col 15 |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 |  |  |  |  |  | Count sheet for Sterile Consumable in  |  |  |  |  |  |  |  |  |  |
| 2 |  |  |  |  |  | Code | Description | Number | Retest/ Expiry | Qty | UoM | Count Qty | Actual Retest/ Expiry | Counted by | Count Date |
| 3 |  |  |  |  |  | 0020 | ACETIC ACID - GLACIAL | 260018 |  | 2470 | mL |  |  |  |  |
| 4 | Count Sheet Selection |  |  |  |  | 0195 | BENZALKONIUM CHLORIDE SOLUTION 50% | 260013 |  | 0 | mL |  |  |  |  |
| 5 | RM Category: | Sterile Consumable |  |  |  | 0380 | CALCIUM HYDROXIDE | 250248 |  | 0 | g |  |  |  |  |
| 6 | Location: |  |  |  |  | 0630 | CRYSTAL VIOLET C.I. 42555 | 260012 |  | 0 | g |  |  |  |  |
| 7 |  |  |  |  |  | 1160 | METHYL HYDROXYBENZOATE | 260010 |  | 0 | g |  |  |  |  |
| 8 |  |  |  |  |  | 1194 | ORAL-MIX SF | 250250 |  | 0 | mL |  |  |  |  |
| 9 |  |  |  |  |  | 1380 | PROPYLENE GLYCOL | 260011 |  | 0 | mL |  |  |  |  |
| 10 |  |  |  |  |  | 1465 | SIROLIMUS | 260014 |  | 0 | g |  |  |  |  |
| 11 |  |  |  |  |  | 1470 | SODIUM BENZOATE | 250249 |  | 0 | g |  |  |  |  |
| 12 |  |  |  |  |  | 1520 | SODIUM CITRATE | 250247 |  | 0 | g |  |  |  |  |
| 13 |  |  |  |  |  | 1910 | UREA | 260017 |  | 0 | g |  |  |  |  |
| 14 |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| 15 |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |

### Sheet Snapshot: `Retest Profile` (R1C3:R30C15)
| Row | Col 1 | Col 2 | Col 3 | Col 4 | Col 5 | Col 6 | Col 7 | Col 8 | Col 9 | Col 10 | Col 11 | Col 12 | Col 13 |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 |  |  |  |  |  | All Material Retest Date Summary |  |  |  |  |  |  |  |
| 2 |  |  |  |  |  |  |  |  |  |  |  |  |  |
| 3 | Retest Profile Summary |  |  |  |  |  |  |  |  |  |  |  |  |
| 4 |  |  |  |  |  |  |  |  |  |  |  |  |  |
| 5 | Past Due | 1 Month | 3 Month | 6 Month |  |  |  |  |  |  |  |  |  |
| 6 | 3 | 7 | 22 | 27 |  |  |  |  |  |  |  |  |  |
| 7 |  |  |  |  |  |  |  |  |  |  |  |  |  |
| 8 |  |  |  |  |  | Retest Profile Details |  |  |  |  |  |  |  |
| 9 |  |  |  |  |  | Bucket | Code | Description | QC Number | QC Balance | UoM | Upcoming Retest Date | QC Status |
| 10 |  |  |  |  |  | Past Due | 0260 | BETAMETHASONE VALERATE | 220023 | 0.3 | g | 31/01/2026 | RELEASED |
| 11 |  |  |  |  |  | Past Due | 0480 | CHLORBUTOL | 240016 | 84.2 | g | 2/06/2026 | RELEASED |
| 12 |  |  |  |  |  | Past Due | 0885 | COMPOUND HYDROXYBENZOATE SOLUTION APF | 250132 | 672 | mL | 31/01/2026 | SAMPLED |
| 13 |  |  |  |  |  | 1 Month | 0370 | CALCIUM CHLORIDE | 180045 | 116.5 | g | 3/09/2026 | RELEASED |
| 14 |  |  |  |  |  | 1 Month | 0400 | CARNITINE L (LevoCarnitine) | 250034 | 6426 | g | 28/02/2026 | RELEASED |
| 15 |  |  |  |  |  | 1 Month | 0481 | CHLORHEXIDINE ACETATE | 150086 | 2.14 | g | 15/03/2026 | RELEASED |

### Sheet Snapshot: `Expiry Profile` (R1C3:R30C15)
| Row | Col 1 | Col 2 | Col 3 | Col 4 | Col 5 | Col 6 | Col 7 | Col 8 | Col 9 | Col 10 | Col 11 | Col 12 | Col 13 |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 |  |  |  |  |  | Non-ACS Expiry Date Summary |  |  |  |  |  |  |  |
| 2 |  |  |  |  |  |  |  |  |  |  |  |  |  |
| 3 | Expiry Profile Summary |  |  |  |  |  |  |  |  |  |  |  |  |
| 4 |  |  |  |  |  |  |  |  |  |  |  |  |  |
| 5 | Past Due | 1 Month | 3 Month | 6 Month |  |  |  |  |  |  |  |  |  |
| 6 | 3 | 1 | 14 | 12 |  |  |  |  |  |  |  |  |  |
| 7 |  |  |  |  |  |  |  |  |  |  |  |  |  |
| 8 |  |  |  |  |  | Expiry Profile Details |  |  |  |  |  |  |  |
| 9 |  |  |  |  |  | Bucket | Code | Description | QC Number | QC Balance | UoM | Expiry Date | QC Status |
| 10 |  |  |  |  |  | Past Due | 1188 | ORA-SWEET SF | 250165 | 15209 | mL | 30/11/2025 | RELEASED |
| 11 |  |  |  |  |  | Past Due | 2060 | HYPROMELLOSE (METHOPT) 0.5% 15ML EYE DROPS 1 | 250140 | 0 | ea | 31/12/2025 | REJECTED |
| 12 |  |  |  |  |  | Past Due | 2066 | VORICONAZOLE (VFEND) 200MG INJECTION 1 | 250162 | 0 | ea | 31/12/2025 | REJECTED |
| 13 |  |  |  |  |  | 1 Month | 1186 | ORA-PLUS SF | 250166 | 471 | mL | 28/02/2026 | RELEASED |
| 14 |  |  |  |  |  | 3 Month | 1932 | WATER FOR INJECTION 1L | 250151 | 21 | ea | 30/04/2026 | RELEASED |
| 15 |  |  |  |  |  | 3 Month | 2001 | AMIODARONE (CORDARONE X) 200MG TABLETS 30 | 250100 | 18 | ea | 30/04/2026 | RELEASED |

### Sheet Snapshot: `Helper` (R1C2:R30C16)
| Row | Col 1 | Col 2 | Col 3 | Col 4 | Col 5 | Col 6 | Col 7 | Col 8 | Col 9 | Col 10 | Col 11 | Col 12 | Col 13 | Col 14 | Col 15 |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | AdjReason |  | UoMList |  | QCStatus |  | SortQCNumRMCode |  | SortQCNumRMDesc |  | SortBatchNumber |  | SortQCNumber |  | SortFilterQCRMCode |
| 2 | Damaged |  | g |  | Rejected |  |  |  | ACETIC ACID - GLACIAL |  | A010125 |  | 040185 |  | 0020 |
| 3 | Deliverate production overage |  | mL |  | Cancelled |  |  |  | ACETYLCHOLINE CHLORIDE |  | A010225 |  | 050084 |  | 0040 |
| 4 | Inaccurate stock count |  | ea |  | Released |  |  |  | ACETYLCYSTEINE (DBL) 2G/10ML INJECTION |  | A010325 |  | 110121 |  | 0080 |
| 5 | Insufficient content in original container |  |  |  | Quarantine |  |  |  | ACETYLCYSTEINE (DBL) 2G/10ML INJECTION 10 |  | A010425 |  | 110135 |  | 0130 |
| 6 | Missing Stock |  |  |  | Sampled |  |  |  | AGAROSE GEL |  | A010525 |  | 140033 |  | 0131 |
| 7 | Overage in original container |  |  |  | Receipted |  |  |  | AGAROSE POWDER |  | A010525 |  | 140064 |  | 0180 |
| 8 | Production error |  |  |  |  |  |  |  | ALTEPLASE (ACTILYSE) 50MG INJECTION 1 |  | A010625 |  | 150006 |  | 0195 |
| 9 | QC Lab Use |  |  |  |  |  |  |  | ALUMINIUM POTASSIUM SULFATE |  | A010725 |  | 150012 |  | 0260 |
| 10 | QC/Retention Sample |  |  |  |  |  |  |  | AMIKACIN (DBL) 500MG/2ML INJECTION 5 |  | A010825 |  | 150086 |  | 0270 |
| 11 | Spillage |  |  |  |  |  |  |  | AMIODARONE (CORDARONE X) 200MG TABLETS 30 |  | A020125 |  | 150125 |  | 0300 |
| 12 | Stock found/located |  |  |  |  |  |  |  | AMLODIPINE (AMLODIPINE SANDOZ) 5MG TABLETS 30 |  | A020225 |  | 160071 |  | 0310 |
| 13 |  |  |  |  |  |  |  |  | APREPITANT (APREPITANT APOTEX) 165MG CAPSULES 1 |  | A020225 |  | 160106 |  | 0320 |
| 14 |  |  |  |  |  |  |  |  | AQUEOUS CREAM |  | A020325 |  | 170001 |  | 0351 |
| 15 |  |  |  |  |  |  |  |  | ASCORBIC ACID |  | A020325 |  | 170087 |  | 0352 |

### Sheet Snapshot: `QC_Number` (R1C1:R30C15)
| Row | Col 1 | Col 2 | Col 3 | Col 4 | Col 5 | Col 6 | Col 7 | Col 8 | Col 9 | Col 10 | Col 11 | Col 12 | Col 13 | Col 14 | Col 15 |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | QC_Number | RM_Code | Standard_Name | Quick_Code | Brand_Name | Part_ARTG | Vendor_Name | Mfg_Name | Mfg_Lot | Mfg_Expiry | Number_Containers | Total_Qty | Total_Sample_Qty | Master_UoM | GR_Date |
| 2 | 230084 | 1160 | METHYL HYDROXYBENZOATE |  |  |  |  | MEDISCA PLATTSBURGH NY 12901 USA | 190678/A | 31/07/2026 | 1 | 113 |  | g |  |
| 3 | 170087 | 0170 | BALSAM TOLU |  |  |  |  |  |  |  |  |  |  | g |  |
| 4 | 210010 | 1301 | POLIHEXANIDE (PHMB) 20% SOLUTION |  |  |  |  | Fagron St Paul MN 55120 USA | 20A16-F03-368296 |  |  |  |  | mL |  |
| 5 | 140033 | 1270 | PHENOL |  |  |  |  |  |  |  |  |  |  | g |  |
| 6 | 220017 | 0550 | CINNAMALDEHYDE |  |  |  |  | Sigma-Aldrich Macquarie Park NSW 2113 AU | W228613 | 15/03/2026 |  |  | 25 | mL |  |
| 7 | 220023 | 0260 | BETAMETHASONE VALERATE |  |  |  |  | Medisca Irving TX 75063 USA | 179899/E |  |  |  |  | g |  |
| 8 | 220044 | 1861 | TRYPTONE SOYA BROTH |  |  |  |  | Oxoid Ltd Hampshire UK | 3434914 | 28/02/2027 |  |  |  | g |  |
| 9 | 040185 | 0450 | CETOSTEARYL ALCOHOL |  |  |  |  |  |  |  |  |  |  | g |  |
| 10 | 230009 | 1920 | VANILLIN |  |  |  |  | Medisca Plattsburgh NY 12901 USA | 183285/A | 30/06/2026 |  |  | 0.91 | g |  |
| 11 | 180081 | 1310 | POTASSIUM CHLORIDE |  |  |  |  |  |  |  |  |  | 2 | g |  |
| 12 | 230035 | 0590 | CLOVE OIL |  |  |  |  | Sigma-Aldrich St. Louis MO 63103 USA | SHBP8108 | 28/06/2027 |  |  |  | mL |  |
| 13 | 230038 | 0802 | GELATIN CAPSULE - SIZE 0 |  |  |  |  | Medisca Sydney NSW 2020 AU | 805544 | 31/08/2026 |  |  |  | ea |  |
| 14 | 230047 | 0730 | DITHRANOL |  |  |  |  | Medisca Irving TX 75063 USA | C202347 | 28/08/2027 |  |  |  | g |  |
| 15 | 230061 | 0802 | GELATIN CAPSULE - SIZE 0 |  |  |  |  | Medisca Sydney NSW 2020 AU | 805544/E | 31/08/2026 |  |  |  | ea |  |

### Sheet Snapshot: `RM_Master` (R1C1:R30C15)
| Row | Col 1 | Col 2 | Col 3 | Col 4 | Col 5 | Col 6 | Col 7 | Col 8 | Col 9 | Col 10 | Col 11 | Col 12 | Col 13 | Col 14 | Col 15 |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | RM_Code | Std_Name | Alt_Name | Category | Storage_Cond | S4_Number | Quick_Code | Prod_ID | Brand_Name | Part_ARTG | Master_UoM | QTY_Per_Pack | Conc_Vol | Wt_per_mL | Wt_per_Unit |
| 2 | 0010 | ACACIA |  | ACS | General Room Temp |  |  |  |  |  | g |  |  |  |  |
| 3 | 0020 | ACETIC ACID - GLACIAL |  | ACS | Flammable Room Temp |  |  |  |  |  | mL | 2500 |  | 1.049 |  |
| 4 | 0030 | ACETONE |  | ACS | Flammable Room Temp |  |  |  |  |  | mL |  |  | 0.79 |  |
| 5 | 0035 | ACETYLCHOLINE |  | ACS | General Room Temp |  |  |  |  |  | g |  |  |  |  |
| 6 | 0040 | AGAROSE GEL |  | ACS | General Room Temp |  |  |  |  |  | g |  |  |  |  |
| 7 | 0041 | ASPIRIN | ACETYLSALICYLIC ACID | ACS | General Room Temp |  |  |  |  |  | g |  |  |  |  |
| 8 | 0050 | ALANINE | L-ALANINE | ACS | General Room Temp |  |  |  |  |  | g |  |  |  |  |
| 9 | 0060 | ALCOHOLIC SOAP SOLUTION |  | ACS | General Room Temp |  |  |  |  |  | mL |  |  |  |  |
| 10 | 0070 | ALUMINIUM ACETATE SOLUTION |  | ACS | General Room Temp |  |  |  |  |  | mL |  |  |  |  |
| 11 | 0071 | ALUMINIUM HYDROXIDE GEL (DRIED) |  | ACS | General Room Temp |  |  |  |  |  | g |  |  |  |  |
| 12 | 0080 | ALUMINIUM POTASSIUM SULFATE | ALUM - DODECAHYDRATE | ACS | General Room Temp |  |  |  | PCCA | 30-1692-500 | g | 500 |  |  |  |
| 13 | 0090 | AMARANTH C.I. 16185 |  | ACS | General Room Temp |  |  |  |  |  | g |  |  |  |  |
| 14 | 0100 | AMETHOCAINE HYDROCHLORIDE | TETRACAINE HYDROCHLORIDE | ACS | General Room Temp |  |  |  |  |  | g |  |  |  |  |
| 15 | 0101 | AMMONIA SOLUTION - STRONG |  | ACS | General Room Temp |  |  |  |  |  | mL |  |  | 0.91 |  |

### Sheet Snapshot: `Batch_Details` (R1C1:R30C15)
| Row | Col 1 | Col 2 | Col 3 | Col 4 | Col 5 | Col 6 | Col 7 | Col 8 | Col 9 | Col 10 | Col 11 | Col 12 | Col 13 | Col 14 | Col 15 |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | QID_Code | Batch_Number | Manuf_Lot_No | Manuf_Lot_No_2 | Expiry | Quantity | Created_By | Created_Date | Avery_Label_Code | Labels_per_sheet | Total_Printed | Product_ID | CPMU_Ref | Product_Description | Print_Type |
| 2 | Dummy | Dummy | Dummy | Dummy | Dummy | Dummy | Dummy | Dummy | Dummy | Dummy | Dummy | Dummy | Dummy | Dummy | Dummy |
| 3 | 600640.A | G170725 |  |  | 1/07/2026 | 10 | CollettMatt | ######## | L7159 | 24 | 24 | 6450 | ACI-001 | ACETIC ACID 6%   200mL  (LAM) | Original |
| 4 | 601050.B | E120524 |  |  | 1/05/2026 | 10 | CollettMatt | ######## | L7161 | 18 | 18 | 18290 | ACI-001B | ACETIC ACID 3% GREEN   1L  (LAM) | Original |
| 5 | 600515.A | H040625 |  |  | 1/02/2027 | 10 | CollettMatt | ######## | L7159 | 24 | 24 | 6448 | ACI-002 | ACETIC ACID 2%   100mL  (LAM) | Original |
| 6 | 718595.A | A210723 |  |  | 11/08/2023 | 10 | CollettMatt | ######## | L7158 | 30 | 30 | 29389 | APP-004A | FLUCYTOSINE 17% AMPHOTERICIN B 3% VAGINAL CR 5g | Original |
| 7 | 706385.A | F190423 |  |  | 1/12/2024 | 10 | CollettMatt | ######## | L7158 | 30 | 30 | 8469 | CAP-003 | BIOTIN 50mg   60 | Original |
| 8 | 706384.A | F300224 |  |  | 1/01/2026 | 10 | CollettMatt | ######## | L7159 | 24 | 24 | 8468 | CAP-004 | BIOTIN 15mg   60 | Original |
| 9 | 705705.A | E190224 |  |  | 1/11/2025 | 10 | CollettMatt | ######## | L7159 | 24 | 24 | 15062 | CAP-007 | BORIC ACID 600mg   Pessary 15 | Original |
| 10 | 318200.A | H060625 |  |  | 1/10/2025 | 10 | CollettMatt | ######## | L7159 | 24 | 24 | 26143 | CAP-009 | Phenytoin  30mg Capsules 30 New Formulation (LAM) | Original |
| 11 | 318200.A | H060625 | 8169452 |  | 1/10/2025 | 10 | CollettMatt | ######## | L7159 | 24 | 24 | 26143 | CAP-009 | Phenytoin  30mg Capsules 30 New Formulation (LAM) | Original |
| 12 | 318210.A | H060625 | 8169452 |  | 1/10/2025 | 10 | CollettMatt | ######## | L7159 | 24 | 24 | 18210 | CAP-009 | Phenytoin  30mg Capsules 90 New Formulation (LAM) | Original |
| 13 | 318244.A | D100224 | 8184675 |  | 1/07/2026 | 10 | CollettMatt | ######## | L7159 | 24 | 24 | 26145 | CAP-010 | Phenytoin  100mg Capsules 30 New Formulation (LAM) | Original |
| 14 | 8155362.A | H140425 |  |  | 1/12/2025 | 18 | SuL | ######## | L7161 | 18 | 18 | 22504 | SOL-027 | MELATONIN 10mg/mL SUBLINGUAL SOLUTION   20mL | Original |
| 15 | 8155362.A | H140425 |  |  | 1/12/2025 | 396 | SuL | ######## | L7161 | 18 | 396 | 22504 | SOL-027 | MELATONIN 10mg/mL SUBLINGUAL SOLUTION   20mL | Reprint |
