# Compa-Ratio Catch-Up
This tool iterates through an employee roster to spend budget on the lowest paid employees through the compa-ratio catch-up method. Coded in VBA, runs in Excel through a form.

## What is compa-ratio?
In compensation, we measure competitiveness for an employee's compensation with the comparison ratio (compa-ratio for short). Compa-ratio is obtained by dividing the employee's salary by the midpoint of the pay range. We adjust the range accordingly when the employee works less than the number of hours to be considered a full-time employee.

With this, it is common to average the compa-ratios for a whole organization to see how the organization is doing in terms of paying its employees competitively.

## What do we do with compa-ratio?
Occasionally, an HR or compensation professional is presented with budget to adjust underpaid employees. Whenever this happens, a common method to decide how to spend the budget is through getting the lowest compa-ratio and starting to put money there until this person catches up with the next lowest paid employee. This process repeats until all the budget is spent. It can be a very manual process, so I have developed a tool that iterates through the data and distributes the budget accordingly.

This tool has saved me a lot of time so far. It was coded in VBA and is run directly on Excel.<br>
![image](https://github.com/jacksoncaquino/Compa-Ratio-Catch-Up/assets/61064363/4502cee0-9034-48b2-abfc-0715f23225cc)

# How to use those fields?
You need a roster with at least the following columns: <br>
• Compa-Ratio<br>
• Midpoint for the pay range of the job where each employee sits<br>
<br>
1. For the budget, you can provide the budget number by either typing the number to the text edit on the form or by clicking a cell in your spreadsheet that has the total budget. Alternatively, you can also select a range that has the budget for each employee and it will sum it for you.
2. For CR first cell, know that your roster will need to have the employees' current compa ratio. It is important that there are no blank cells for current compa-ratio, or else the whole budget will be spent only on the segment of employees before the first interruption.
3. For midpoints 1st cell, please select the first cell where the midpoints for the employees' pay ranges are.
4. For FTE, you can ignore this if all employees are full-time. If not, you will need a number between 0 and 1 for each employee. For example, if that location has a 40 hour work week and you a an employee who works 20 hours a week, FTE should be 0.5. Alternatively, if you have already adjusted the pay ranges to that, you do not need to include this FTE information.
5. For eligibility, you are able to indicate those who should not receive a raise even if their Compa-ratio is low. For example, you may determine that newly hired employees or employees who are low performers should not receive a raise. If that's the case, you may flag those to the tool by adding the word "No" to that column. Any blanks, "Yes", or anything other than a "No" will be eligible for a raise.
6. Spend is where the tool will place how much of your budget will go to that employee. Select only a blank cell for the first employee on the column where you want the values placed.

# Installing this tool in your roster file:
After downloading the two files (FRM and FRX) from this repository, if you need assistance importing the FRM file to your Excel file, follow the instructions below:
1. On Excel, press alt + F11 to open the Visual Basic Editor
2. On the Visual Basic editor, right-click your file and then click on "import file":<br>
![image](https://github.com/jacksoncaquino/Compa-Ratio-Catch-Up/assets/61064363/a8632f0a-0c0b-4fb6-b759-a5ca4f32cbd1)

3. Choose your FRM file that you downloaded from this repository
4. You'll now have the form installed in your file
5. To show the form, you'll need to create a macro
6. Right-click the file again, hover over "insert" and then click "module":<br>
![image](https://github.com/jacksoncaquino/Compa-Ratio-Catch-Up/assets/61064363/0fd1d498-0bf7-451d-bf07-f9a502d6e768)
7. On the newly created module, type the following:
sub cr_catchup()
  CR_Catchup_Form_addin.Show
endsub
8. Close the Visual Basic editor and go back to your Excel file
9. Go to View, Macros, look for a macro called cr_catchup
10. Click the macro name and then click run:<br>
![image](https://github.com/jacksoncaquino/Compa-Ratio-Catch-Up/assets/61064363/586c27b4-a4e0-47b2-a900-3ca934adfbf4)
11. The form should show up and you can now use it.


