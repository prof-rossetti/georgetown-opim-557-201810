# Retirement Savings Calculator

> Adapted from a deliverable created by Professor Dillon-Merrill.

Assume you own and operate a financial planning business which helps customers plan for retirement. Your objective is to build yourself a tool to automate the common calculations required to provide your clients with retirement savings advice.

## Learning Objectives

  1. Design and build a tool to perform automated calculations and aid a decision-making process.
  2. Find practical applications for practicing new programming concepts, primarily loops.

## Instructions

Create a new macro-enabled workbook named "`netid`-simple-system.xlsm", where `netid` is your university-issued net identifier (i.e. the first part of your university-issued email address).

Rename the first sheet to "Interface".

Your submission should adhere to the following requirements, as detailed in the corresponding sections below:

  + [Information Requirements](#information-requirements)
  + [User Interface and Experience (UI/UX) Requirements](#ui--ux-requirements)
  + [Validation Requirements](#validation-requirements)
  + [Calculation Requirements](#calculation-requirements)

### Information Requirements

Regardless of how you choose to capture and display inputs and outputs, make sure the user sees only properly-formatted values. Rates should be formatted with a percent sign (`%`) and dollar amounts should be formatted as USD with a dollar sign (`$`) and two decimal places.

#### Inputs

Your system should accept the following user inputs:

  1. The client's **current age**.
  2. The client's **desired retirement age**.
  3. The client's current amount of savings (i.e. **initial savings balance**). Assume the client does not have any debt.
  4. The amount of money the client plans to contribute to savings each year (i.e. **annual contribution**). Assume contributions are made at the end of each year, after interest has been accrued.
  5. A projected annual growth rate for the client's savings (i.e **annual interest rate**). Assume interest will compound on an annual basis (at the end of each year), not on a monthly basis.

Example Inputs:

![a screenshot showing a message box displaying the user inputs](/img/projects/savings-calc/display-inputs.png)

The table below provides a framework for you to translate these information inputs into VBA variables:

Information Input | Example Variable Name | Example Variable Datatype | Example Value
--- | ---  | ---  | ---
Current Age | `Age` | `Integer` | `60`
Desired Retirement Age | `RetirementAge` | `Integer` | `65`
Initial Savings Balance | `InitialBalance` | `Double` | `50000.00`
Annual Savings Contribution | `AnnualContribution` | `Double` | `18000.00`
Annual Savings Growth Rate (Interest Rate) | `InterestRate` | `Double` | `0.05`

> NOTE: Depending on your interface, when declaring variables, it may be reasonable to declare different datatypes than those suggested above, or to not declare datatypes at all. Use your best judgement and go with what works.

#### Outputs

Your system should produce the following outputs:

  1. The **final savings balance** at the end of the year when the client reaches the specified retirement age.
  2. The portion of the final savings balance which was contributed directly by the client (i.e. **total savings contribution**).
  3. The portion of the final savings balance resulting from accrued interest on the principal (i.e. **total interest accrued**).

Example Outputs:

![a screenshot showing a message box displaying the final outputs](/img/projects/savings-calc/display-outputs.png)

The table below provides a framework for you to translate these information inputs into VBA variables:

Information Output | Example Variable Name | Example Datatype | Example Value
--- | ---  | ---  | ---
Final Savings Balance | `SavingsBalance` | `Double` | `189439.21`
Total Savings Contribution | `TotalContribution` | `Double` | `158000.00`
Total Interest Accrued | `AccruedInterest` | `Double` | `31439.21`

> NOTE: Depending on your interface, when declaring variables, it may be reasonable to declare different datatypes than those suggested above, or to not declare datatypes at all. Use your best judgement and go with what works.

See the "Calculation Requirements" section below for more information about how to calculate these information outputs.

### UI/UX Requirements

Provide written instructions which explain how to use the tool.

Use any kind of interface you'd like (e.g. cells, input boxes, ActiveX Controls, user forms, etc.) to capture user inputs, as long as it is appropriate. You may draw inspiration from any of these [example interfaces](/projects/savings-calc/example-interfaces.md). But beware, some interface elements may be more appropriate than others, and your job is to choose the interface elements that will provide the best user experience. NOTE: If you end up using input boxes, make sure you handle situations where the user clicks "Cancel".

Include a button control that when clicked will: read and validate the inputs, perform the calculations, and produce the outputs. Outputs should also be properly formatted (see above).

If inputs and outputs are ever visible at the same time, they should always correspond with eachother. In other words, previously-generated outputs should not be visible at the same time as yet-to-be used inputs. Practically, this means you should clear output values if the user starts to adjust any of the input values.

The user should never experience runtime errors or be prompted to "debug" the code.

### Validation Requirements

Do your best to prevent the user from inputting values which are invalid (e.g. entering a value of the wrong data type, entering a value outside of a reasonable range of accepted values, etc.).

If the program doesn't completely prevent the user from submitting invalid inputs, the program must validate those inputs. If the program detects an invalid user input, it should immediately stop execution and display a friendly message to the user describing what went wrong and how the user can fix the problem.

### Calculation Requirements

The figure below depicts an example of the system's desired calculations. It is NOT meant to depict the user interface, the way the system captures inputs, or the way the system produces outputs. It is also NOT meant to depict the manner in which the system performs calculations, because the calculations should be performed using VBA, not cell formulae. For clarification, the program does not need to write cell values like this. Again, this is just an example to help you test to make sure your calculations are producing the expected results.

![a screenshot depicting some example inputs, example annual calculations, and example final outputs](/img/projects/savings-calc/example-calculation-results.png)

#### Annual Calculations

The savings balance at the end of any given year is equal to the initial savings balance for that year, plus the amount of accrued interest for that year, plus the end-of-year contribution.

The amount of accrued interest for any given year is equal to the initial savings balance for that year times the annual interest rate. Note: the end-of-year contribution does not accrue interest during the year it is contributed.

The initial savings balance for any given year is equal to the ending savings balance from the previous year.

#### Calculating Final Outputs

The final savings balance is the balance at the end of the year when the client hits the desired retirement age. For example, if the desired retirement age is 65, the program should output the savings balance at the end of that year, after that year's interest accrual and savings contribution are calculated (see "Annual Calculations" above).

The final amount of accrued interest is equal to the sum of interest accrued during each year between the user's current age and desired retirement age, inclusive.

The final amount of contributed principal is equal to the initial savings balance plus the sum of all end-of-year contributions, which is also equal to the final savings balance less the final amount of accrued interest.








## Submission Instructions

Upload your workbook file [to Canvas](https://georgetown.instructure.com/courses/65741/assignments/165668).

## Evaluation Methodology

Submissions will be evaluated based on their ability to meet each of the component requirements (see corresponding sections above):

Requirements Category | System Requirement | Weight
-- | -- | --
Information | Captures inputs. | 0.10
Information | Displays final outputs. | 0.10
Information | Formats inputs and outputs (USD, pct, etc.). | 0.10
UI/UX | Provides written user instructions. | 0.08
UI/UX | Calculations triggered by button-click event. | 0.05
UI/UX | Reasonable user experience, with instructional clarity, lacking runtime errors. | 0.12
Validation | Validates age inputs. | 0.07
Validation | Validates age less than retirement age. | 0.05
Validation | Validates currency inputs. | 0.07
Validation | Validates percentage input. | 0.06
Calculation | Calculates final outputs with accuracy. | 0.20

This rubric is tentative, and may be subject to slight adjustments during the grading process.

Additionally, the professor reserves the right to award extra credit for submissions which exceed expectations, and/or submissions which demonstrate successful implementation of each "Further Exploration Challenge" (see below).

<hr>

## Further Exploration Challenges (Optional)

> WARNING: This challenge is optional. Only attempt this challenge if/once you have successfully completed all other basic project requirements. Prefer to submit a project which perfectly meets basic requirements over a project which attempts to address this challenge but fails to perfectly meet all basic requirements.

### Writing Annual Data to Sheet

Challenge TBA

### Probabilistic Annual Returns

Challenge TBA

### Charts and Graphs

Challenge TBA
