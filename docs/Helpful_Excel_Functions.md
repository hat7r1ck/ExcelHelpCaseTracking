# Helpful Excel Functions and Formulas

This document provides a collection of Excel formulas and functions that can be used in the Excel Help Case Tracking workbook. These formulas assist with data lookup, time calculations, conditional logic, and performance analysis. You can reference and adapt these formulas to meet the specific requirements of your project.

---

## 1. VLOOKUP for Data Preview

Use this formula on the **QuickEntry** sheet to preview related data from the **Data_Import** sheet (formatted as a table, e.g., `XSOARData`).

**Example:**

```excel
=IFERROR(VLOOKUP(B2, XSOARData, 2, FALSE), "Not Found")
```

- **B2:** The cell where the CaseID is entered.
- **XSOARData:** The table containing your imported case data.
- **2:** Returns data from the second column (e.g., TimeCreated).

---

## 2. INDEX/MATCH Alternative for Lookup

For more flexibility than VLOOKUP—especially if columns might be rearranged—use the INDEX/MATCH combination.

**Example:**

```excel
=IFERROR(INDEX(XSOARData[TimeCreated], MATCH(B2, XSOARData[CaseID], 0)), "Not Found")
```

- **XSOARData[TimeCreated]:** The column with the case creation timestamps.
- **XSOARData[CaseID]:** The column containing the case IDs.
- **B2:** The lookup value (CaseID).

---

## 3. Calculating Time Differences in Minutes

To calculate the delay between two timestamps (for example, when a case is picked up versus when it was received), use the following formula. Multiplying by 1440 converts the time difference from days to minutes.

**Example:**

```excel
=(TimePickedUp - TimeReceived) * 1440
```

- **TimePickedUp:** A cell reference to when the case was picked up.
- **TimeReceived:** A cell reference to when the case was received.

---

## 4. COUNTIFS for Spike Detection

To detect a spike in case alerts (by counting the number of alerts within a specified time window), you can use the COUNTIFS function.

**Example:**

```excel
=COUNTIFS(Alerts[AlertReceivedTime], ">=" & (B2 - TIME(0,30,0)), Alerts[AlertReceivedTime], "<" & B2)
```

- **Alerts[AlertReceivedTime]:** The column that records when each alert is received.
- **B2:** The reference time (for example, the current case’s pickup time).
- **TIME(0,30,0):** Specifies a 30-minute interval.

---

## 5. Average Response Time Calculation

To determine the average response time (in minutes) between when a case is created and when help is logged, use the following formula.

**Example:**

```excel
=AVERAGE(HelpCaseLog[HelpTimestamp] - HelpCaseLog[TimeCreated]) * 1440
```

- **HelpCaseLog[HelpTimestamp]:** The column with the help timestamps.
- **HelpCaseLog[TimeCreated]:** The column with the case creation timestamps.
- Multiplying by 1440 converts the average time difference from days to minutes.

---

## 6. IF Function for Conditional Logic

Apply conditional logic to label data. For example, to label a delay as "Long Delay" if it is 30 minutes or more:

**Example:**

```excel
=IF(A2 >= 30, "Long Delay", "Normal")
```

- **A2:** A cell reference containing the calculated delay in minutes.
- Returns "Long Delay" if the value is 30 or more; otherwise, returns "Normal."
