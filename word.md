# Microsoft Word Tips and Tricks

# Customize mailing list fields [Shift + F9 OR Alt + F9]
- number
  ```
  {MERGEFIELD Number\#,0}
  \#0       for 1000
  \#,0      for 1,000
  \#,0.00   for 1,000.00
  ```
- currency
  ```
  {MERGEFIELD Currency\#$,0}
  \#$0       for $1000
  \#$,0      for $1,000
  \#$,0.00   for $1,000.00
  \#"$,0.00;($,0.00);'-'"   for ($3,000.00) negative value, - [hyphen] for zero
  ```
- percentage
  ```
  {MERGEFIELD Percent\#0.00%}
  \#0.00%       for 100.00%
  \#0%      for 100%
  ```
- date and time
  ```
  {MERGEFIELD Date \@ "dddd, MMMM dd, yyyy"}
  \@"M/d/yyyy"                      for     5/20/2020
  \@"d-MMM-yy"                      for     20-May-22
  \@"d MMMM yyyy"                   for     20 May 2022
  \@"ddd, d MMMM yyyy"              for     Fri, 20 May 2022
  \@"dddd, d MMMM yyyy"             for     Friday, 20 May 2022
  \@"dddd, MMMM dd, yyyy"           for     Friday, May 20, 2022
  \@"h:mm AM/PM"                    for     10:45 PM
  \@"HH:mm"                         for     22:45
  \@"HH:mm:ss"                      for     22:45:30
  ```

# 