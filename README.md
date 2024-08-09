# Cohort Data
This project was for a very specific use case in combination with a Google Sheet. Below is the setup of the sheets.

## Form Responses
|Timestamp|Who's Entering the data?|Company Name|Cohort Year|Year|Revenue|Expenses|New Company Name|Has the company been funded?|Amount of funding|
|---|---|---|---|---|---|---|---|---|---|
|`Date`|`String`|`String`|`Integer`|`Integer`|`Float`|`Float`|`String`|`Boolean`|`Float`|

## Raw Data
| Company | Cohort Year | Year | Revenue | Expenses | Difference | Company ID | Cleaned Company | Funded |
|---|---|---|---|---|---|---|---|---|
| `String` | `Integer` | `Integer` | `Float` | `Float` | `Float` | `String` | `String` | `Boolean` |

## Edited Data
| Company | Cohort Year | Year | Revenue | Expenses | Difference | Company ID | Cleaned Company | Funded | Timestamp |
|---|---|---|---|---|---|---|---|---|---|
| `String` | `Integer` | `Integer` | `Float` | `Float` | `Float` | `String` | `String` | `Boolean` | `Date` |

## Company List
| Company | Active | Cohort Years | Funded | Funding Amount |
|---|---|---|---|---|
| `String` | `Boolean` | `Integer` or `List of Integers` | `Boolean` | `Float` |

## Relative Year Summary
| Company ID | Company | Cohort Year | Year 1 |  | Year 2 |  | Year 3 |  | Year 4 |  |
|---|---|---|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
|  | Carry Over: | `Boolean` | Difference | Percentage | Difference | Percentage | Difference | Percentage | Difference | Percentage |
| `String` | `String` | `Integer` | `Float` | `Float` | `Float` | `Float` | `Float` | `Float` | `Float` | `Float` |

## Company Growth Dashboard
| 1 |  |
|---|---|
| 2 | Funded |
| 3 | `Boolean` |
| 4 | Non-Funded |
| 5 | `Boolean` |
| 6 |  |
| 7 | Year 1 |
| 8 | `Boolean` |
| 9 | Year 2 |
| 10 | `Boolean` |
| 11 | Year 3 |
| 12 | `Boolean` |
| 13 | Year 4 |
| 14 | `Boolean` |
| 15 |  |
| 16 | Carry Over from Previous Year |
| 17 | `Boolean` |
| 18 |  |
| 19 | Sort By |
| 20 | `Dropdown of Strings` |
| 21 |  |
| 22 | Include Companies with No Data |
| 23 | `Boolean` |
| 24 |  |

Leave `D28:X` empty. It will be modified with the changes in the above cells.

## Budget Filter
| 1 | Funded |
|---|:---:|
| 2 | `Boolean` |
| 3 | Non-Funded |
| 4 | `Boolean` |
| 5 |  |
| 6 | All Years |
| 7 | `Boolean` |
| 8 | Most Recent |
| 9 | `Boolean` |
| 10 |  |
| 11 | Budget |
| 12 | `Dropdown of Strings` |
| 13 | Sort By |
| 14 | `Dropdown of Strings` |
| 15 | Sort By Column |
| 16 | `Dropdown of Strings` |
| 17 |  |
| 18 | Display Column |
| 19 | `Dropdown of Strings` |
| 20 |  |
| 21 | Include Companies with No Data |
| 22 | `Boolean` |
| 23 |  |

Leave `D29:P` empty. It will be modified with the changes in the above cells.

## Individual Company Dashboard
Have a Dropdown (from a range) with the source as `='Company List'!$A$2:$A`.

In another part in the same sheet, use the following formula:
`=QUERY('Raw Data'!A1:F, "SELECT * WHERE A = '" & A2 & "'")`.

Note that in this example, `A2` is where the dropdown is stored.

## Companies With Missing Data
| 1 | Refresh |  |  |  |  |  |
|---|:---:|---|---|---|---|---|
| 2 | `Boolean` |  |  |  |  |  |
| 3 |  | Funded |  |  | Non-Funded |  |
| 4 | Separate Funded and Non-Funded | Company | Years |  | Company | Years |
| 5 | `Boolean` | `String` | `Integer` |  | `String` | `Integer` or `List of Integers` |