"use strict";
const LESSONS = [
  {
    "heading": "Adding prices",
    "explanation": "When a customer buys two products, their final cost is the first price plus the second price. In Excel, + joins the values from two cells; it does not join their text labels.",
    "method": "Read the amount in A2. Read the amount in B2. Add them together and place the formula in C2.",
    "reflect": "Change either product price. The total should update without editing the formula."
  },
  {
    "heading": "Money left over",
    "explanation": "A budget is money available before spending. Remaining money is the starting budget minus expenses, so the order of the subtraction matters.",
    "method": "Start with A2 (budget). Subtract B2 (expenses). A negative result means overspending, not a broken formula.",
    "reflect": "If expenses become larger than the budget, what sign should the result have?"
  },
  {
    "heading": "Price multiplied by quantity",
    "explanation": "Buying several identical items costs the price of one item multiplied by how many you buy. Excel uses an asterisk (*) for multiplication.",
    "method": "Take the unit price in A2. Multiply by the count in B2. The result in C2 is the total purchase cost.",
    "reflect": "If quantity becomes zero, what should the total be?"
  },
  {
    "heading": "Sharing a cost equally",
    "explanation": "To split a bill evenly, divide the entire bill by the number of people. The slash (/) is the division operator in Excel.",
    "method": "A2 holds the complete bill. B2 holds the number of people. Divide A2 by B2; do not reverse them.",
    "reflect": "What happens if the number of people is zero?"
  },
  {
    "heading": "Calculate in the right order",
    "explanation": "Excel normally multiplies before it subtracts. Parentheses tell Excel to calculate the discounted unit price first, then multiply by quantity.",
    "method": "Calculate A2 minus B2 inside parentheses. Multiply that amount by D2, the number of items.",
    "reflect": "Compare (50−10)×3 with 50−10×3. Why are the answers different?"
  },
  {
    "heading": "Adding a whole range",
    "explanation": "A range such as A2:A6 means every cell from A2 through A6. SUM adds the numeric values in that group, so you do not need a long chain of plus signs.",
    "method": "Identify the first and last expense cells. Put a colon between them in SUM. Include every day, not only the endpoints.",
    "reflect": "If an extra day is inserted inside the range, what should happen to the total?"
  },
  {
    "heading": "Formulas that follow the row",
    "explanation": "A relative reference such as A2 points to a cell relative to the formula. When you copy a formula down, Excel changes its row number automatically.",
    "method": "For the first row, multiply A2 by B2. When copied to the next row, the intended inputs are A3 and B3.",
    "reflect": "Why would $A$2 be a bad choice for the price of each separate row?"
  },
  {
    "heading": "Keep one tax cell fixed",
    "explanation": "Sometimes all rows must use a single shared value, such as a VAT rate in F1. Dollar signs lock that cell address so a copied formula keeps using it.",
    "method": "Use the price from the current row. Multiply by one plus the VAT rate from $F$1. When copied down, the price changes but $F$1 does not.",
    "reflect": "Which reference should change when the formula moves from row 2 to row 3?"
  },
  {
    "heading": "Spot a wrong cell reference",
    "explanation": "Excel can return a number even when a formula points to the wrong cell. This is a logical error, which is often harder to notice than an Excel error message.",
    "method": "Read the row being calculated. Check that the price and quantity are both from row 2. Replace the unrelated B3 reference.",
    "reflect": "Why is a plausible-looking result not proof that a formula is correct?"
  },
  {
    "heading": "Build a small invoice",
    "explanation": "An invoice combines several rules: each item has a price and quantity, the line totals are added, and tax is applied to the subtotal.",
    "method": "Calculate each line total. Add those totals using SUM. Apply the shared tax rate and calculate the final invoice amount.",
    "reflect": "If one item quantity changes, which invoice figures should recalculate?"
  },
  {
    "heading": "Sum several entries",
    "explanation": "SUM is a function that adds numeric cells in a range. It is easier to maintain than typing =A2+A3+A4 and so on.",
    "method": "Select the first and last cost cells. Use SUM with the full range in parentheses.",
    "reflect": "If one daily cost changes, will SUM recalculate?"
  },
  {
    "heading": "Mean versus total",
    "explanation": "AVERAGE finds the arithmetic mean: add the numeric scores and divide by how many numeric scores there are. It is not the same as SUM.",
    "method": "Identify the score range. Apply AVERAGE once to that range rather than dividing by a guessed fixed count.",
    "reflect": "What happens to the average if a new very high score is added?"
  },
  {
    "heading": "Finding the biggest value",
    "explanation": "MAX returns the largest numeric value in a group. It does not add numbers or tell you which row the maximum came from.",
    "method": "Select the sales cells. Call MAX on that range and compare the result with each individual sale.",
    "reflect": "If the largest sale is replaced by a smaller number, what changes?"
  },
  {
    "heading": "Finding the smallest value",
    "explanation": "MIN returns the smallest numeric entry in a range. Negative numbers can be the minimum too.",
    "method": "Select the temperature cells and return the lowest observed temperature.",
    "reflect": "Would −3 or 2 be the minimum?"
  },
  {
    "heading": "Counting numbers only",
    "explanation": "COUNT counts cells containing numerical values. Text labels and blank cells are not counted as numbers.",
    "method": "Inspect the mixed dataset and count only entries that are actually numeric.",
    "reflect": "If a number is imported as text, will COUNT necessarily count it?"
  },
  {
    "heading": "Counting all filled cells",
    "explanation": "COUNTA counts cells that are not empty, including text and numbers. That makes it useful for checking submitted responses.",
    "method": "Identify the response range and count entries that contain a value. Compare with COUNT to see the difference.",
    "reflect": "Will the name “Alex” contribute to COUNTA?"
  },
  {
    "heading": "Rounding to decimal places",
    "explanation": "ROUND changes a number to the specified number of decimal places using Excel rounding rules. For currency, two decimal places are common.",
    "method": "Give ROUND the value and the number 2 for decimal places. Check values immediately above and below the rounding boundary.",
    "reflect": "Does ROUND change the underlying input cell or only the calculated output?"
  },
  {
    "heading": "Always round away from zero",
    "explanation": "ROUNDUP moves a nonzero number away from zero to the selected precision. It is useful for charging whole shipping units.",
    "method": "Apply ROUNDUP to the shipping weight with zero decimal places to obtain whole units.",
    "reflect": "Why can 2.1 units require 3 whole units?"
  },
  {
    "heading": "Always round toward zero",
    "explanation": "ROUNDDOWN truncates a number toward zero at the specified precision. This is not the same as ordinary rounding.",
    "method": "Choose the numeric value and the required number of decimal places. Check how the result behaves for negatives as well.",
    "reflect": "How is ROUNDDOWN different from ROUND when the removed digit is 9?"
  },
  {
    "heading": "Summarise scores in several ways",
    "explanation": "A useful dashboard needs different questions answered from one dataset: total, average, highest, and lowest. Each function has a distinct job.",
    "method": "Complete every requested output cell with the correct summary function and the same source range.",
    "reflect": "Why is one large formula not always the clearest way to present a dashboard?"
  },
  {
    "heading": "Percent means out of 100",
    "explanation": "A percentage is a share of a whole. When a percentage is stored as the number 15, divide by 100 to convert it to 0.15.",
    "method": "Multiply the base amount by the percent value, then divide by 100.",
    "reflect": "What is 10% of 200, and why is it not 2000?"
  },
  {
    "heading": "Discount reduces a price",
    "explanation": "A discount removes a percentage of the original price. A 20% discount means the customer pays 80%, not 20%.",
    "method": "Compute one minus the discount fraction and multiply by the original price.",
    "reflect": "Would a 0% discount change the price?"
  },
  {
    "heading": "Percentage salary increase",
    "explanation": "An increase adds part of the original amount. A 10% raise means the new salary is 110% of the old salary.",
    "method": "Multiply the salary by (1 + percentage/100).",
    "reflect": "What would a 0% increase do?"
  },
  {
    "heading": "VAT as a separate amount",
    "explanation": "A VAT amount is calculated as a share of the pre-tax price. It is different from the final VAT-inclusive price.",
    "method": "Multiply the net price by the VAT rate, using the correct representation of the rate.",
    "reflect": "If the pre-tax price doubles, what happens to the VAT amount?"
  },
  {
    "heading": "Price including VAT",
    "explanation": "An inclusive total combines the original amount and its tax. One plus the tax fraction gives the multiplier.",
    "method": "Take the original price and multiply it by (1 + rate).",
    "reflect": "What is the difference between VAT amount and VAT-inclusive amount?"
  },
  {
    "heading": "Growth compared with the starting value",
    "explanation": "Percentage change measures the difference relative to the original value, not the new one. The usual calculation is (new−old)/old.",
    "method": "Subtract the previous value from the new value, then divide by the previous value.",
    "reflect": "Why can division by zero be a problem when the original value is zero?"
  },
  {
    "heading": "Profit is revenue minus cost",
    "explanation": "Gross profit in this exercise is the selling price minus the cost price. It is an amount of money, not a percentage.",
    "method": "Identify revenue and cost. Subtract cost from revenue, keeping the direction correct.",
    "reflect": "What does negative gross profit mean?"
  },
  {
    "heading": "Margin measures profit relative to selling price",
    "explanation": "Profit margin compares profit with revenue. A common error is to divide by cost, which gives markup instead.",
    "method": "Find the profit (revenue−cost) and divide it by revenue. Convert to a percentage if required.",
    "reflect": "Why are margin and markup different percentages?"
  },
  {
    "heading": "Growth that compounds",
    "explanation": "Compound growth means each period starts from the previous period’s increased amount. Exponents calculate repeated multiplication efficiently.",
    "method": "Use the initial amount, a growth multiplier (1+r), and raise it to the number of periods.",
    "reflect": "Why is growth over 3 years not always 3 times the first-year increase?"
  },
  {
    "heading": "Pricing with a business rule",
    "explanation": "Sometimes a discount is allowed only when a condition is met. IF selects between discounted and full price.",
    "method": "Check the condition first, then calculate the discount only for the qualifying case.",
    "reflect": "What should happen exactly at the discount threshold?"
  },
  {
    "heading": "A test that returns pass or fail",
    "explanation": "IF chooses between two outputs after evaluating a TRUE/FALSE condition. The comparison must match the school’s pass threshold.",
    "method": "Compare score with passing mark using >= when equality counts. Put the required labels in quotation marks.",
    "reflect": "Should a score equal to the pass mark pass?"
  },
  {
    "heading": "Classify working hours",
    "explanation": "A threshold condition such as “at least 40” must include 40 itself. In Excel that is >=, not >.",
    "method": "Read hours and required minimum from their cells. Return the correct text for TRUE and FALSE.",
    "reflect": "What changes if you use > instead of >=?"
  },
  {
    "heading": "Reorder below the limit",
    "explanation": "A stock alert can be triggered only when stock falls strictly below a limit. The word “below” means <, not <=.",
    "method": "Compare available quantity with reorder point, then choose the two message strings.",
    "reflect": "Should quantity exactly equal to the reorder point trigger an alert?"
  },
  {
    "heading": "Two requirements at once",
    "explanation": "AND is TRUE only if every listed condition is TRUE. Use it when two exam passes are both required.",
    "method": "Build each pass comparison and join them inside AND, then use IF for the label.",
    "reflect": "If one score fails, is AND still TRUE?"
  },
  {
    "heading": "Either condition is enough",
    "explanation": "OR is TRUE when at least one condition is TRUE. It is different from AND, which needs all conditions.",
    "method": "Write the two acceptable qualification checks. Wrap them with OR and choose the result with IF.",
    "reflect": "If both conditions are false, what does OR return?"
  },
  {
    "heading": "Reverse a logical result",
    "explanation": "NOT turns TRUE into FALSE and FALSE into TRUE. It is useful when a rule is easier to describe negatively.",
    "method": "Identify the original condition and wrap it in NOT. Decide which output corresponds to TRUE.",
    "reflect": "What is NOT(TRUE)?"
  },
  {
    "heading": "Multiple grade bands",
    "explanation": "Nested IF lets one decision lead to another. Test higher grade boundaries first so a high score is not classified as a lower grade.",
    "method": "Evaluate the top score threshold. If it fails, evaluate the next grade and so on.",
    "reflect": "Why is the order of grade conditions important?"
  },
  {
    "heading": "Multiple alternatives with IFS",
    "explanation": "IFS checks conditions in order and returns the value associated with the first TRUE condition. A final TRUE condition can supply a default.",
    "method": "List each grade boundary and its label in order. Include a fall-through outcome.",
    "reflect": "What happens when none of the IFS tests is TRUE and no default is provided?"
  },
  {
    "heading": "Fix the equality edge case",
    "explanation": "Logic bugs often hide at the boundary: greater than is not the same as greater than or equal to.",
    "method": "Test one below, exactly equal to, and one above the boundary. Correct the comparison that fails equality.",
    "reflect": "Which single input is most useful for catching this exact bug?"
  },
  {
    "heading": "Eligibility with several criteria",
    "explanation": "A bonus should be awarded only when performance AND attendance meet their minimums. IF translates that combined condition into a label.",
    "method": "Write two >= comparisons, combine with AND, then return Bonus or No Bonus.",
    "reflect": "Would OR incorrectly grant a bonus when only one target is met?"
  },
  {
    "heading": "Count records matching one rule",
    "explanation": "COUNTIF looks at a range and counts cells meeting one criterion. The criterion >=50 is written as a text condition.",
    "method": "Choose the score range and the quoted comparison string.",
    "reflect": "Will 50 itself count when the criterion is >=50?"
  },
  {
    "heading": "Count values under a threshold",
    "explanation": "COUNTIF can identify stock quantities below a threshold. Less-than excludes the boundary itself.",
    "method": "Select the quantity range. Supply a criterion such as \"<10\" and check exactly 10.",
    "reflect": "Does a stock quantity of 10 match <10?"
  },
  {
    "heading": "Sum just one category",
    "explanation": "SUMIF has two jobs: find rows meeting a criterion and sum the matching values in another range. Both ranges must align.",
    "method": "Use the region range for matching East; use the sales range for values to add.",
    "reflect": "Why would summing the region text column not work?"
  },
  {
    "heading": "Average one group",
    "explanation": "AVERAGEIF calculates the mean of values only for rows matching one criterion.",
    "method": "Identify the class labels, choose the class, and average the corresponding score cells.",
    "reflect": "Should a student in another class affect this average?"
  },
  {
    "heading": "Count an intersection",
    "explanation": "COUNTIFS counts rows where every criterion is satisfied. Each criteria range should refer to the same set of records.",
    "method": "Pair each criteria range with its condition: region and payment status.",
    "reflect": "Does East plus Pending count as East plus Paid?"
  },
  {
    "heading": "Sum with several filters",
    "explanation": "SUMIFS first takes the numerical sum range, then pairs of criteria ranges and conditions. It differs from SUMIF’s argument order.",
    "method": "Identify sales as the amount to total, then filter by region and product.",
    "reflect": "What is the first argument in SUMIFS?"
  },
  {
    "heading": "Average with several filters",
    "explanation": "AVERAGEIFS averages values when all criteria agree. One mismatching criterion excludes the row.",
    "method": "Choose the salaries to average and apply both department and contract-type criteria.",
    "reflect": "Does an IT worker with the wrong contract type qualify?"
  },
  {
    "heading": "Match the start of a word",
    "explanation": "The * wildcard can represent any number of characters. A criterion such as Pro* counts text beginning with Pro.",
    "method": "Use COUNTIF with a quoted wildcard criterion, not multiplication.",
    "reflect": "Would Project and Profile both match Pro*?"
  },
  {
    "heading": "Use a threshold from a cell",
    "explanation": "Hardcoded thresholds can become stale. Concatenate a comparison operator with a cell reference using & to build a dynamic criterion.",
    "method": "Build the text \">=\" and join it to F1. COUNTIF then tests against the current F1 value.",
    "reflect": "If F1 increases, should the formula need editing?"
  },
  {
    "heading": "Build a regional reporting summary",
    "explanation": "Different business questions require different conditional functions: one for counts, one for totals, and one for totals with multiple filters.",
    "method": "Calculate East orders, East sales value, and East Paid sales value in their specified cells.",
    "reflect": "Why can COUNTIF not directly provide the total sales value?"
  },
  {
    "heading": "Recover from division errors",
    "explanation": "IFERROR tries a calculation first. If it produces an Excel error, a fallback message is returned instead.",
    "method": "Place the division in the first argument and the message in the second.",
    "reflect": "Should valid zero results be replaced by the fallback?"
  },
  {
    "heading": "Recognise text input",
    "explanation": "ISTEXT returns TRUE for text data and FALSE for non-text values. It helps identify values imported with the wrong type.",
    "method": "Point ISTEXT at the input cell and inspect whether it is actually text.",
    "reflect": "Does the number 100 behave the same as the text \"100\" for every calculation?"
  },
  {
    "heading": "Recognise a truly empty cell",
    "explanation": "ISBLANK checks whether a referenced cell is empty. A cell holding a formula that returns an empty string is not necessarily blank.",
    "method": "Use ISBLANK on the required input and distinguish a truly empty cell from visible text.",
    "reflect": "Is a cell containing a space genuinely empty?"
  },
  {
    "heading": "Give a useful fallback message",
    "explanation": "Error handling improves usability when users enter invalid data. IFERROR should not silently hide errors that need investigation.",
    "method": "Wrap the calculation and provide the exact user-facing fallback required.",
    "reflect": "Why is “Not available” sometimes better than displaying #DIV/0!?"
  },
  {
    "heading": "Catch missing data specifically",
    "explanation": "NA() produces the #N/A error used for unavailable values. IFNA catches only #N/A, unlike IFERROR, which catches a wider set of errors.",
    "method": "First make the appropriate missing-data condition return NA(). Wrap it with IFNA to show Missing.",
    "reflect": "Will IFNA catch a divide-by-zero error?"
  },
  {
    "heading": "Check for a number",
    "explanation": "ISNUMBER distinguishes numerical values from labels and numeric-looking text. It can validate imported fields before calculations.",
    "method": "Evaluate the value in A2 with ISNUMBER.",
    "reflect": "Will a text cell holding \"45\" necessarily pass ISNUMBER?"
  },
  {
    "heading": "Combine validation with a message",
    "explanation": "IF can use a type-check function as its condition. This produces a clear message instead of a TRUE/FALSE result.",
    "method": "Put ISNUMBER(A2) inside IF. Return Numeric when true and Check input otherwise.",
    "reflect": "Which output should appear for text data?"
  },
  {
    "heading": "Wrong row, believable result",
    "explanation": "A misplaced reference may return a normal number from a different row. The correct fix is to identify the intended record, not just make the formula run.",
    "method": "Compare the row of the numerator with the row of the divisor, and correct the misplaced reference.",
    "reflect": "Why might Excel show no error for a reference to B3?"
  },
  {
    "heading": "Ask whether an expression fails",
    "explanation": "ISERROR tests whether a calculation would produce an error and returns TRUE or FALSE. It does not replace the error with another value.",
    "method": "Place the risky division inside ISERROR. Test both zero and nonzero divisors.",
    "reflect": "How is ISERROR different from IFERROR?"
  },
  {
    "heading": "Report both input quality and errors",
    "explanation": "A robust report answers three separate questions: Is the input numeric? What is a safe displayed result? Does the calculation error?",
    "method": "Use a type check, an error-handled calculation, and an error-status calculation in the three output cells.",
    "reflect": "Why is it useful to report an error flag even when a friendly fallback is displayed?"
  }
];
