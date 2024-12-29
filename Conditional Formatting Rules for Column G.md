Here are the conditional formatting formulas for column G, structured elegantly in Markdown for clarity and reuse:

---

## **Conditional Formatting Rules for Column G**

### 1. **Two Fields Populated**
- **Formula:**  
  `=COUNTA(B2:D2)=2`
- **Range:**  
  `G2:G1000`
- **Format:**  
  Background: Orange

---

### 2. **One Field Populated**
- **Formula:**  
  `=COUNTA(B2:D2)=1`
- **Range:**  
  `G2:G1000`
- **Format:**  
  Background: Red

---

### 3. **Row Marked Complete**
- **Formula:**  
  `=G2="Complete"`
- **Range:**  
  `G2:G1000`
- **Format:**  
  Background: Green

---

### 4. **Missing Multiple Fields**
- **Formula:**  
  `=AND(ISNUMBER(SEARCH("Missing", G2)))`
- **Range:**  
  `G2:G1000`
- **Format:**  
  Background: Yellow

---

### 5. **Partial Completion (Between 1 and 3 Fields Populated)**
- **Formula:**  
  `=AND(COUNTA(A2:D2)>0, COUNTA(A2:D2)<4)`
- **Range:**  
  `G2:G1000`
- **Format:**  
  Background: Light Yellow

---

To apply these rules:
1. Highlight the range `G2:G1000` in your Google Sheet.
2. Navigate to **Format > Conditional Formatting**.
3. Add each rule using the **Custom Formula Is** option.
4. Assign the specified format to each rule.
5. Save your changes.

This configuration ensures comprehensive and visually distinct conditional formatting for all possible cases in column G.