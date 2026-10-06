/* Ledger — function catalogue for Insert Function, Function Arguments and the formula ScreenTips.
 * Format per line: NAME|arg1;arg2;[optional];more...|Description */
(function (root) {
  'use strict';
  const L = root.L;
  const FI = (L.fninfo = {});
  const DATA = {
    Financial: `
ACCRINT|issue;first_interest;settlement;rate;par;frequency;[basis];[calc_method]|Returns the accrued interest for a security that pays periodic interest.
ACCRINTM|issue;settlement;rate;par;[basis]|Returns the accrued interest for a security that pays interest at maturity.
AMORDEGRC|cost;date_purchased;first_period;salvage;period;rate;[basis]|Returns the prorated linear depreciation of an asset for each accounting period, using a depreciation coefficient.
AMORLINC|cost;date_purchased;first_period;salvage;period;rate;[basis]|Returns the prorated linear depreciation of an asset for each accounting period.
COUPDAYBS|settlement;maturity;frequency;[basis]|Returns the number of days from the beginning of the coupon period to the settlement date.
COUPDAYS|settlement;maturity;frequency;[basis]|Returns the number of days in the coupon period that contains the settlement date.
COUPDAYSNC|settlement;maturity;frequency;[basis]|Returns the number of days from the settlement date to the next coupon date.
COUPNCD|settlement;maturity;frequency;[basis]|Returns the next coupon date after the settlement date.
COUPNUM|settlement;maturity;frequency;[basis]|Returns the number of coupons payable between the settlement date and maturity date.
COUPPCD|settlement;maturity;frequency;[basis]|Returns the previous coupon date before the settlement date.
CUMIPMT|rate;nper;pv;start_period;end_period;type|Returns the cumulative interest paid between two periods.
CUMPRINC|rate;nper;pv;start_period;end_period;type|Returns the cumulative principal paid on a loan between two periods.
DB|cost;salvage;life;period;[month]|Returns the depreciation of an asset for a specified period using the fixed-declining balance method.
DDB|cost;salvage;life;period;[factor]|Returns the depreciation of an asset for a specified period using the double-declining balance method or some other method you specify.
DISC|settlement;maturity;pr;redemption;[basis]|Returns the discount rate for a security.
DOLLARDE|fractional_dollar;fraction|Converts a dollar price, expressed as a fraction, into a dollar price, expressed as a decimal number.
DOLLARFR|decimal_dollar;fraction|Converts a dollar price, expressed as a decimal number, into a dollar price, expressed as a fraction.
DURATION|settlement;maturity;coupon;yld;frequency;[basis]|Returns the annual duration of a security with periodic interest payments.
EFFECT|nominal_rate;npery|Returns the effective annual interest rate.
FV|rate;nper;pmt;[pv];[type]|Returns the future value of an investment based on periodic, constant payments and a constant interest rate.
FVSCHEDULE|principal;schedule|Returns the future value of an initial principal after applying a series of compound interest rates.
INTRATE|settlement;maturity;investment;redemption;[basis]|Returns the interest rate for a fully invested security.
IPMT|rate;per;nper;pv;[fv];[type]|Returns the interest payment for a given period for an investment, based on periodic, constant payments and a constant interest rate.
IRR|values;[guess]|Returns the internal rate of return for a series of cash flows.
ISPMT|rate;per;nper;pv|Returns the interest paid during a specific period of an investment.
MDURATION|settlement;maturity;coupon;yld;frequency;[basis]|Returns the Macauley modified duration for a security with an assumed par value of $100.
MIRR|values;finance_rate;reinvest_rate|Returns the internal rate of return for a series of periodic cash flows, considering both cost of investment and interest on reinvestment of cash.
NOMINAL|effect_rate;npery|Returns the annual nominal interest rate.
NPER|rate;pmt;pv;[fv];[type]|Returns the number of periods for an investment based on periodic, constant payments and a constant interest rate.
NPV|rate;value1;[value2];...|Returns the net present value of an investment based on a discount rate and a series of future payments (negative values) and income (positive values).
ODDFPRICE|settlement;maturity;issue;first_coupon;rate;yld;redemption;frequency;[basis]|Returns the price per $100 face value of a security with an odd first period.
ODDFYIELD|settlement;maturity;issue;first_coupon;rate;pr;redemption;frequency;[basis]|Returns the yield of a security with an odd first period.
ODDLPRICE|settlement;maturity;last_interest;rate;yld;redemption;frequency;[basis]|Returns the price per $100 face value of a security with an odd last period.
ODDLYIELD|settlement;maturity;last_interest;rate;pr;redemption;frequency;[basis]|Returns the yield of a security with an odd last period.
PDURATION|rate;pv;fv|Returns the number of periods required by an investment to reach a specified value.
PMT|rate;nper;pv;[fv];[type]|Calculates the payment for a loan based on constant payments and a constant interest rate.
PPMT|rate;per;nper;pv;[fv];[type]|Returns the payment on the principal for a given investment based on periodic, constant payments and a constant interest rate.
PRICE|settlement;maturity;rate;yld;redemption;frequency;[basis]|Returns the price per $100 face value of a security that pays periodic interest.
PRICEDISC|settlement;maturity;discount;redemption;[basis]|Returns the price per $100 face value of a discounted security.
PRICEMAT|settlement;maturity;issue;rate;yld;[basis]|Returns the price per $100 face value of a security that pays interest at maturity.
PV|rate;nper;pmt;[fv];[type]|Returns the present value of an investment: the total amount that a series of future payments is worth now.
RATE|nper;pmt;pv;[fv];[type];[guess]|Returns the interest rate per period of a loan or an investment.
RECEIVED|settlement;maturity;investment;discount;[basis]|Returns the amount received at maturity for a fully invested security.
RRI|nper;pv;fv|Returns an equivalent interest rate for the growth of an investment.
SLN|cost;salvage;life|Returns the straight-line depreciation of an asset for one period.
SYD|cost;salvage;life;per|Returns the sum-of-years' digits depreciation of an asset for a specified period.
TBILLEQ|settlement;maturity;discount|Returns the bond-equivalent yield for a Treasury bill.
TBILLPRICE|settlement;maturity;discount|Returns the price per $100 face value for a Treasury bill.
TBILLYIELD|settlement;maturity;pr|Returns the yield for a Treasury bill.
VDB|cost;salvage;life;start_period;end_period;[factor];[no_switch]|Returns the depreciation of an asset for any period you specify, including partial periods, using the double-declining balance method or some other method you specify.
XIRR|values;dates;[guess]|Returns the internal rate of return for a schedule of cash flows that is not necessarily periodic.
XNPV|rate;values;dates|Returns the net present value for a schedule of cash flows that is not necessarily periodic.
YIELD|settlement;maturity;rate;pr;redemption;frequency;[basis]|Returns the yield on a security that pays periodic interest.
YIELDDISC|settlement;maturity;pr;redemption;[basis]|Returns the annual yield for a discounted security, for example a Treasury bill.
YIELDMAT|settlement;maturity;issue;rate;pr;[basis]|Returns the annual yield of a security that pays interest at maturity.`,
    'Date & Time': `
DATE|year;month;day|Returns the number that represents the date in Ledger date-time code.
DATEDIF|start_date;end_date;unit|Calculates the number of days, months or years between two dates.
DATEVALUE|date_text|Converts a date in the form of text to a number that represents the date in Ledger date-time code.
DAY|serial_number|Returns the day of the month, a number from 1 to 31.
DAYS|end_date;start_date|Returns the number of days between the two dates.
DAYS360|start_date;end_date;[method]|Returns the number of days between two dates based on a 360-day year (twelve 30-day months).
EDATE|start_date;months|Returns the serial number of the date that is the indicated number of months before or after the start date.
EOMONTH|start_date;months|Returns the serial number of the last day of the month before or after a specified number of months.
HOUR|serial_number|Returns the hour as a number from 0 (12:00 A.M.) to 23 (11:00 P.M.).
ISOWEEKNUM|date|Returns the ISO week number in the year for a given date.
MINUTE|serial_number|Returns the minute, a number from 0 to 59.
MONTH|serial_number|Returns the month, a number from 1 (January) to 12 (December).
NETWORKDAYS|start_date;end_date;[holidays]|Returns the number of whole workdays between two dates.
NETWORKDAYS.INTL|start_date;end_date;[weekend];[holidays]|Returns the number of whole workdays between two dates with custom weekend parameters.
NOW||Returns the current date and time formatted as a date and time.
SECOND|serial_number|Returns the second, a number from 0 to 59.
TIME|hour;minute;second|Converts hours, minutes and seconds given as numbers to a Ledger serial number, formatted with a time format.
TIMEVALUE|time_text|Converts a text time to a Ledger serial number for a time, a number from 0 (12:00:00 AM) to 0.999988426 (11:59:59 PM).
TODAY||Returns the current date formatted as a date.
WEEKDAY|serial_number;[return_type]|Returns a number from 1 to 7 identifying the day of the week of a date.
WEEKNUM|serial_number;[return_type]|Returns the week number in the year.
WORKDAY|start_date;days;[holidays]|Returns the serial number of the date before or after a specified number of workdays.
WORKDAY.INTL|start_date;days;[weekend];[holidays]|Returns the serial number of the date before or after a specified number of workdays with custom weekend parameters.
YEAR|serial_number|Returns the year of a date, an integer in the range 1900 - 9999.
YEARFRAC|start_date;end_date;[basis]|Returns the year fraction representing the number of whole days between start_date and end_date.`,
    'Math & Trig': `
ABS|number|Returns the absolute value of a number, a number without its sign.
ACOS|number|Returns the arccosine of a number, in radians in the range 0 to Pi.
ACOSH|number|Returns the inverse hyperbolic cosine of a number.
ACOT|number|Returns the arccotangent of a number, in radians in the range 0 to Pi.
ACOTH|number|Returns the inverse hyperbolic cotangent of a number.
AGGREGATE|function_num;options;ref1;...|Returns an aggregate in a list or database, optionally ignoring hidden rows and error values.
ARABIC|text|Converts a Roman numeral to Arabic.
ASIN|number|Returns the arcsine of a number in radians, in the range -Pi/2 to Pi/2.
ASINH|number|Returns the inverse hyperbolic sine of a number.
ATAN|number|Returns the arctangent of a number in radians, in the range -Pi/2 to Pi/2.
ATAN2|x_num;y_num|Returns the arctangent of the specified x- and y- coordinates, in radians between -Pi and Pi, excluding -Pi.
ATANH|number|Returns the inverse hyperbolic tangent of a number.
BASE|number;radix;[min_length]|Converts a number into a text representation with the given radix (base).
CEILING|number;significance|Rounds a number up, to the nearest multiple of significance.
CEILING.MATH|number;[significance];[mode]|Rounds a number up, to the nearest integer or to the nearest multiple of significance.
COMBIN|number;number_chosen|Returns the number of combinations for a given number of items.
COMBINA|number;number_chosen|Returns the number of combinations with repetitions for a given number of items.
COS|number|Returns the cosine of an angle.
COSH|number|Returns the hyperbolic cosine of a number.
COT|number|Returns the cotangent of an angle.
COTH|number|Returns the hyperbolic cotangent of a number.
CSC|number|Returns the cosecant of an angle.
CSCH|number|Returns the hyperbolic cosecant of an angle.
DECIMAL|number;radix|Converts a text representation of a number in a given base into a decimal number.
DEGREES|angle|Converts radians to degrees.
EVEN|number|Rounds a positive number up and negative number down to the nearest even integer.
EXP|number|Returns e raised to the power of a given number.
FACT|number|Returns the factorial of a number, equal to 1*2*3*...* Number.
FACTDOUBLE|number|Returns the double factorial of a number.
FLOOR|number;significance|Rounds a number down to the nearest multiple of significance.
FLOOR.MATH|number;[significance];[mode]|Rounds a number down, to the nearest integer or to the nearest multiple of significance.
GCD|number1;[number2];...|Returns the greatest common divisor.
INT|number|Rounds a number down to the nearest integer.
LCM|number1;[number2];...|Returns the least common multiple.
LN|number|Returns the natural logarithm of a number.
LOG|number;[base]|Returns the logarithm of a number to the base you specify.
LOG10|number|Returns the base-10 logarithm of a number.
MDETERM|array|Returns the matrix determinant of an array.
MINVERSE|array|Returns the inverse matrix for the matrix stored in an array.
MMULT|array1;array2|Returns the matrix product of two arrays, an array with the same number of rows as array1 and columns as array2.
MOD|number;divisor|Returns the remainder after a number is divided by a divisor.
MROUND|number;multiple|Returns a number rounded to the desired multiple.
MULTINOMIAL|number1;[number2];...|Returns the multinomial of a set of numbers.
MUNIT|dimension|Returns the unit matrix for the specified dimension.
ODD|number|Rounds a positive number up and negative number down to the nearest odd integer.
PI||Returns the value of Pi, 3.14159265358979, accurate to 15 digits.
POWER|number;power|Returns the result of a number raised to a power.
PRODUCT|number1;[number2];...|Multiplies all the numbers given as arguments.
QUOTIENT|numerator;denominator|Returns the integer portion of a division.
RADIANS|angle|Converts degrees to radians.
RAND||Returns a random number greater than or equal to 0 and less than 1, evenly distributed (changes on recalculation).
RANDARRAY|[rows];[columns];[min];[max];[integer]|Returns an array of random numbers.
RANDBETWEEN|bottom;top|Returns a random number between the numbers you specify.
ROMAN|number;[form]|Converts an Arabic numeral to Roman, as text.
ROUND|number;num_digits|Rounds a number to a specified number of digits.
ROUNDDOWN|number;num_digits|Rounds a number down, toward zero.
ROUNDUP|number;num_digits|Rounds a number up, away from zero.
SEC|number|Returns the secant of an angle.
SECH|number|Returns the hyperbolic secant of an angle.
SEQUENCE|rows;[columns];[start];[step]|Returns a sequence of numbers.
SERIESSUM|x;n;m;coefficients|Returns the sum of a power series based on the formula.
SIGN|number|Returns the sign of a number: 1 if the number is positive, zero if the number is zero, or -1 if the number is negative.
SIN|number|Returns the sine of an angle.
SINH|number|Returns the hyperbolic sine of a number.
SQRT|number|Returns the square root of a number.
SQRTPI|number|Returns the square root of (number * Pi).
SUBTOTAL|function_num;ref1;...|Returns a subtotal in a list or database.
SUM|number1;[number2];...|Adds all the numbers in a range of cells.
SUMIF|range;criteria;[sum_range]|Adds the cells specified by a given condition or criteria.
SUMIFS|sum_range;criteria_range1;criteria1;...|Adds the cells specified by a given set of conditions or criteria.
SUMPRODUCT|array1;[array2];...|Returns the sum of the products of corresponding ranges or arrays.
SUMSQ|number1;[number2];...|Returns the sum of the squares of the arguments. The arguments can be numbers, arrays, names or references to cells that contain numbers.
SUMX2MY2|array_x;array_y|Sums the differences between the squares of two corresponding ranges or arrays.
SUMX2PY2|array_x;array_y|Returns the sum total of the sums of squares of numbers in two corresponding ranges or arrays.
SUMXMY2|array_x;array_y|Sums the squares of the differences in two corresponding ranges or arrays.
TAN|number|Returns the tangent of an angle.
TANH|number|Returns the hyperbolic tangent of a number.
TRUNC|number;[num_digits]|Truncates a number to an integer by removing the decimal, or fractional, part of the number.`,
    Statistical: `
AVEDEV|number1;[number2];...|Returns the average of the absolute deviations of data points from their mean.
AVERAGE|number1;[number2];...|Returns the average (arithmetic mean) of its arguments, which can be numbers or names, arrays or references that contain numbers.
AVERAGEA|value1;[value2];...|Returns the average (arithmetic mean) of its arguments, evaluating text and FALSE in arguments as 0; TRUE evaluates as 1.
AVERAGEIF|range;criteria;[average_range]|Finds average (arithmetic mean) for the cells specified by a given condition or criteria.
AVERAGEIFS|average_range;criteria_range1;criteria1;...|Finds average (arithmetic mean) for the cells specified by a given set of conditions or criteria.
BETA.DIST|x;alpha;beta;cumulative;[A];[B]|Returns the beta probability distribution function.
BETA.INV|probability;alpha;beta;[A];[B]|Returns the inverse of the cumulative beta probability density function.
BETADIST|x;alpha;beta;[A];[B]|Returns the cumulative beta probability density function.
BETAINV|probability;alpha;beta;[A];[B]|Returns the inverse of the cumulative beta probability density function.
BINOM.DIST|number_s;trials;probability_s;cumulative|Returns the individual term binomial distribution probability.
BINOM.DIST.RANGE|trials;probability_s;number_s;[number_s2]|Returns the probability of a trial result using a binomial distribution.
BINOM.INV|trials;probability_s;alpha|Returns the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value.
BINOMDIST|number_s;trials;probability_s;cumulative|Returns the individual term binomial distribution probability.
CHIDIST|x;deg_freedom|Returns the one-tailed probability of the chi-squared distribution.
CHIINV|probability;deg_freedom|Returns the inverse of the one-tailed probability of the chi-squared distribution.
CHISQ.DIST|x;deg_freedom;cumulative|Returns the left-tailed probability of the chi-squared distribution.
CHISQ.DIST.RT|x;deg_freedom|Returns the right-tailed probability of the chi-squared distribution.
CHISQ.INV|probability;deg_freedom|Returns the inverse of the left-tailed probability of the chi-squared distribution.
CHISQ.INV.RT|probability;deg_freedom|Returns the inverse of the right-tailed probability of the chi-squared distribution.
CHISQ.TEST|actual_range;expected_range|Returns the test for independence.
CHITEST|actual_range;expected_range|Returns the test for independence: the value from the chi-squared distribution for the statistic and the appropriate degrees of freedom.
CONFIDENCE|alpha;standard_dev;size|Returns the confidence interval for a population mean.
CONFIDENCE.NORM|alpha;standard_dev;size|Returns the confidence interval for a population mean, using a normal distribution.
CONFIDENCE.T|alpha;standard_dev;size|Returns the confidence interval for a population mean, using a Student's T distribution.
CORREL|array1;array2|Returns the correlation coefficient between two data sets.
COUNT|value1;[value2];...|Counts the number of cells in a range that contain numbers.
COUNTA|value1;[value2];...|Counts the number of cells in a range that are not empty.
COUNTBLANK|range|Counts the number of empty cells in a specified range of cells.
COUNTIF|range;criteria|Counts the number of cells within a range that meet the given condition.
COUNTIFS|criteria_range1;criteria1;...|Counts the number of cells specified by a given set of conditions or criteria.
COVAR|array1;array2|Returns covariance, the average of the products of deviations for each data point pair in two data sets.
COVARIANCE.P|array1;array2|Returns population covariance, the average of the products of deviations for each data point pair in two data sets.
COVARIANCE.S|array1;array2|Returns the sample covariance, the average of the products of deviations for each data point pair in two data sets.
CRITBINOM|trials;probability_s;alpha|Returns the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value.
DEVSQ|number1;[number2];...|Returns the sum of squares of deviations of data points from their sample mean.
EXPON.DIST|x;lambda;cumulative|Returns the exponential distribution.
EXPONDIST|x;lambda;cumulative|Returns the exponential distribution.
F.DIST|x;deg_freedom1;deg_freedom2;cumulative|Returns the (left-tailed) F probability distribution (degree of diversity) for two data sets.
F.DIST.RT|x;deg_freedom1;deg_freedom2|Returns the (right-tailed) F probability distribution (degree of diversity) for two data sets.
F.INV|probability;deg_freedom1;deg_freedom2|Returns the inverse of the (left-tailed) F probability distribution.
F.INV.RT|probability;deg_freedom1;deg_freedom2|Returns the inverse of the (right-tailed) F probability distribution.
F.TEST|array1;array2|Returns the result of an F-test, the two-tailed probability that the variances in Array1 and Array2 are not significantly different.
FDIST|x;deg_freedom1;deg_freedom2|Returns the F probability distribution (degree of diversity) for two data sets.
FINV|probability;deg_freedom1;deg_freedom2|Returns the inverse of the F probability distribution.
FISHER|x|Returns the Fisher transformation.
FISHERINV|y|Returns the inverse of the Fisher transformation.
FORECAST|x;known_y's;known_x's|Calculates, or predicts, a future value along a linear trend by using existing values.
FORECAST.LINEAR|x;known_y's;known_x's|Calculates, or predicts, a future value along a linear trend by using existing values.
FREQUENCY|data_array;bins_array|Calculates how often values occur within a range of values and then returns a vertical array of numbers having one more element than Bins_array.
FTEST|array1;array2|Returns the result of an F-test, the two-tailed probability that the variances in Array1 and Array2 are not significantly different.
GAMMA|x|Returns the Gamma function value.
GAMMA.DIST|x;alpha;beta;cumulative|Returns the gamma distribution.
GAMMA.INV|probability;alpha;beta|Returns the inverse of the gamma cumulative distribution.
GAMMADIST|x;alpha;beta;cumulative|Returns the gamma distribution.
GAMMAINV|probability;alpha;beta|Returns the inverse of the gamma cumulative distribution.
GAMMALN|x|Returns the natural logarithm of the gamma function.
GAMMALN.PRECISE|x|Returns the natural logarithm of the gamma function.
GAUSS|x|Returns 0.5 less than the standard normal cumulative distribution.
GEOMEAN|number1;[number2];...|Returns the geometric mean of an array or range of positive numeric data.
GROWTH|known_y's;[known_x's];[new_x's];[const]|Returns numbers in an exponential growth trend matching known data points.
HARMEAN|number1;[number2];...|Returns the harmonic mean of a data set of positive numbers: the reciprocal of the arithmetic mean of reciprocals.
HYPGEOM.DIST|sample_s;number_sample;population_s;number_pop;cumulative|Returns the hypergeometric distribution.
HYPGEOMDIST|sample_s;number_sample;population_s;number_pop|Returns the hypergeometric distribution.
INTERCEPT|known_y's;known_x's|Calculates the point at which a line will intersect the y-axis by using a best-fit regression line plotted through the known x-values and y-values.
KURT|number1;[number2];...|Returns the kurtosis of a data set.
LARGE|array;k|Returns the k-th largest value in a data set. For example, the fifth largest number.
LINEST|known_y's;[known_x's];[const];[stats]|Returns statistics that describe a linear trend matching known data points, by fitting a straight line using the least squares method.
LOGEST|known_y's;[known_x's];[const];[stats]|Returns statistics that describe an exponential curve matching known data points.
LOGINV|probability;mean;standard_dev|Returns the inverse of the lognormal cumulative distribution function of x, where ln(x) is normally distributed with parameters Mean and Standard_dev.
LOGNORM.DIST|x;mean;standard_dev;cumulative|Returns the lognormal distribution of x, where ln(x) is normally distributed with parameters Mean and Standard_dev.
LOGNORM.INV|probability;mean;standard_dev|Returns the inverse of the lognormal cumulative distribution function of x.
LOGNORMDIST|x;mean;standard_dev|Returns the cumulative lognormal distribution of x, where ln(x) is normally distributed with parameters Mean and Standard_dev.
MAX|number1;[number2];...|Returns the largest value in a set of values. Ignores logical values and text.
MAXA|value1;[value2];...|Returns the largest value in a set of values. Does not ignore logical values and text.
MAXIFS|max_range;criteria_range1;criteria1;...|Returns the maximum value among cells specified by a given set of conditions or criteria.
MEDIAN|number1;[number2];...|Returns the median, or the number in the middle of the set of given numbers.
MIN|number1;[number2];...|Returns the smallest number in a set of values. Ignores logical values and text.
MINA|value1;[value2];...|Returns the smallest value in a set of values. Does not ignore logical values and text.
MINIFS|min_range;criteria_range1;criteria1;...|Returns the minimum value among cells specified by a given set of conditions or criteria.
MODE|number1;[number2];...|Returns the most frequently occurring, or repetitive, value in an array or range of data.
MODE.MULT|number1;[number2];...|Returns a vertical array of the most frequently occurring, or repetitive, values in an array or range of data.
MODE.SNGL|number1;[number2];...|Returns the most frequently occurring, or repetitive, value in an array or range of data.
NEGBINOM.DIST|number_f;number_s;probability_s;cumulative|Returns the negative binomial distribution.
NEGBINOMDIST|number_f;number_s;probability_s|Returns the negative binomial distribution, the probability that there will be Number_f failures before the Number_s-th success.
NORM.DIST|x;mean;standard_dev;cumulative|Returns the normal distribution for the specified mean and standard deviation.
NORM.INV|probability;mean;standard_dev|Returns the inverse of the normal cumulative distribution for the specified mean and standard deviation.
NORM.S.DIST|z;cumulative|Returns the standard normal distribution (has a mean of zero and a standard deviation of one).
NORM.S.INV|probability|Returns the inverse of the standard normal cumulative distribution.
NORMDIST|x;mean;standard_dev;cumulative|Returns the normal cumulative distribution for the specified mean and standard deviation.
NORMINV|probability;mean;standard_dev|Returns the inverse of the normal cumulative distribution for the specified mean and standard deviation.
NORMSDIST|z|Returns the standard normal cumulative distribution (has a mean of zero and a standard deviation of one).
NORMSINV|probability|Returns the inverse of the standard normal cumulative distribution (has a mean of zero and a standard deviation of one).
PEARSON|array1;array2|Returns the Pearson product moment correlation coefficient, r.
PERCENTILE|array;k|Returns the k-th percentile of values in a range.
PERCENTILE.EXC|array;k|Returns the k-th percentile of values in a range, where k is in the range 0..1, exclusive.
PERCENTILE.INC|array;k|Returns the k-th percentile of values in a range, where k is in the range 0..1, inclusive.
PERCENTRANK|array;x;[significance]|Returns the rank of a value in a data set as a percentage of the data set.
PERCENTRANK.EXC|array;x;[significance]|Returns the rank of a value in a data set as a percentage of the data set (0..1, exclusive).
PERCENTRANK.INC|array;x;[significance]|Returns the rank of a value in a data set as a percentage of the data set (0..1, inclusive).
PERMUT|number;number_chosen|Returns the number of permutations for a given number of objects that can be selected from the total objects.
PERMUTATIONA|number;number_chosen|Returns the number of permutations for a given number of objects (with repetitions) that can be selected from the total objects.
PHI|x|Returns the value of the density function for a standard normal distribution.
POISSON|x;mean;cumulative|Returns the Poisson distribution.
POISSON.DIST|x;mean;cumulative|Returns the Poisson distribution.
PROB|x_range;prob_range;lower_limit;[upper_limit]|Returns the probability that values in a range are between two limits or equal to a lower limit.
QUARTILE|array;quart|Returns the quartile of a data set.
QUARTILE.EXC|array;quart|Returns the quartile of a data set, based on percentile values from 0..1, exclusive.
QUARTILE.INC|array;quart|Returns the quartile of a data set, based on percentile values from 0..1, inclusive.
RANK|number;ref;[order]|Returns the rank of a number in a list of numbers: its size relative to other values in the list.
RANK.AVG|number;ref;[order]|Returns the rank of a number in a list of numbers; if more than one value has the same rank, the average rank is returned.
RANK.EQ|number;ref;[order]|Returns the rank of a number in a list of numbers; if more than one value has the same rank, the top rank of that set of values is returned.
RSQ|known_y's;known_x's|Returns the square of the Pearson product moment correlation coefficient through the given data points.
SKEW|number1;[number2];...|Returns the skewness of a distribution: a characterization of the degree of asymmetry of a distribution around its mean.
SKEW.P|number1;[number2];...|Returns the skewness of a distribution based on a population.
SLOPE|known_y's;known_x's|Returns the slope of the linear regression line through the given data points.
SMALL|array;k|Returns the k-th smallest value in a data set. For example, the fifth smallest number.
STANDARDIZE|x;mean;standard_dev|Returns a normalized value from a distribution characterized by a mean and standard deviation.
STDEV|number1;[number2];...|Estimates standard deviation based on a sample (ignores logical values and text in the sample).
STDEV.P|number1;[number2];...|Calculates standard deviation based on the entire population given as arguments (ignores logical values and text).
STDEV.S|number1;[number2];...|Estimates standard deviation based on a sample (ignores logical values and text in the sample).
STDEVA|value1;[value2];...|Estimates standard deviation based on a sample, including logical values and text. Text and FALSE have the value 0; TRUE has the value 1.
STDEVP|number1;[number2];...|Calculates standard deviation based on the entire population given as arguments (ignores logical values and text).
STDEVPA|value1;[value2];...|Calculates standard deviation based on an entire population, including logical values and text.
STEYX|known_y's;known_x's|Returns the standard error of the predicted y-value for each x in a regression.
T.DIST|x;deg_freedom;cumulative|Returns the left-tailed Student's t-distribution.
T.DIST.2T|x;deg_freedom|Returns the two-tailed Student's t-distribution.
T.DIST.RT|x;deg_freedom|Returns the right-tailed Student's t-distribution.
T.INV|probability;deg_freedom|Returns the left-tailed inverse of the Student's t-distribution.
T.INV.2T|probability;deg_freedom|Returns the two-tailed inverse of the Student's t-distribution.
T.TEST|array1;array2;tails;type|Returns the probability associated with a Student's t-Test.
TDIST|x;deg_freedom;tails|Returns the Student's t-distribution.
TINV|probability;deg_freedom|Returns the two-tailed inverse of the Student's t-distribution.
TREND|known_y's;[known_x's];[new_x's];[const]|Returns numbers in a linear trend matching known data points, using the least squares method.
TRIMMEAN|array;percent|Returns the mean of the interior portion of a set of data values.
TTEST|array1;array2;tails;type|Returns the probability associated with a Student's t-Test.
VAR|number1;[number2];...|Estimates variance based on a sample (ignores logical values and text in the sample).
VAR.P|number1;[number2];...|Calculates variance based on the entire population (ignores logical values and text in the population).
VAR.S|number1;[number2];...|Estimates variance based on a sample (ignores logical values and text in the sample).
VARA|value1;[value2];...|Estimates variance based on a sample, including logical values and text. Text and FALSE have the value 0; TRUE has the value 1.
VARP|number1;[number2];...|Calculates variance based on the entire population (ignores logical values and text in the population).
VARPA|value1;[value2];...|Calculates variance based on the entire population, including logical values and text.
WEIBULL|x;alpha;beta;cumulative|Returns the Weibull distribution.
WEIBULL.DIST|x;alpha;beta;cumulative|Returns the Weibull distribution.
Z.TEST|array;x;[sigma]|Returns the one-tailed P-value of a z-test.
ZTEST|array;x;[sigma]|Returns the one-tailed P-value of a z-test.`,
    'Lookup & Reference': `
ADDRESS|row_num;column_num;[abs_num];[a1];[sheet_text]|Creates a cell reference as text, given specified row and column numbers.
AREAS|reference|Returns the number of areas in a reference. An area is a range of contiguous cells or a single cell.
CHOOSE|index_num;value1;[value2];...|Chooses a value or action to perform from a list of values, based on an index number.
CHOOSECOLS|array;col_num1;...|Returns the specified columns from an array.
CHOOSEROWS|array;row_num1;...|Returns the specified rows from an array.
COLUMN|[reference]|Returns the column number of a reference.
COLUMNS|array|Returns the number of columns in an array or reference.
DROP|array;rows;[columns]|Drops rows or columns from the start or end of an array.
EXPAND|array;rows;[columns];[pad_with]|Expands an array to the specified dimensions.
FILTER|array;include;[if_empty]|Filters a range or array.
FORMULATEXT|reference|Returns a formula as a string.
GETPIVOTDATA|data_field;pivot_table;[field1];[item1];...|Extracts data stored in a PivotTable.
HLOOKUP|lookup_value;table_array;row_index_num;[range_lookup]|Looks for a value in the top row of a table or array of values and returns the value in the same column from a row you specify.
HSTACK|array1;[array2];...|Appends arrays horizontally and in sequence.
HYPERLINK|link_location;[friendly_name]|Creates a shortcut or jump that opens a document stored on your hard drive, a network server, or on the Internet.
INDEX|array;row_num;[column_num];[area_num]|Returns a value or reference of the cell at the intersection of a particular row and column, in a given range.
INDIRECT|ref_text;[a1]|Returns the reference specified by a text string.
LOOKUP|lookup_value;lookup_vector;[result_vector]|Looks up a value either from a one-row or one-column range or from an array.
MATCH|lookup_value;lookup_array;[match_type]|Returns the relative position of an item in an array that matches a specified value in a specified order.
OFFSET|reference;rows;cols;[height];[width]|Returns a reference to a range that is a given number of rows and columns from a given reference.
ROW|[reference]|Returns the row number of a reference.
ROWS|array|Returns the number of rows in a reference or an array.
RTD|progID;server;topic1;...|Retrieves real-time data from a program that supports COM automation.
SORT|array;[sort_index];[sort_order];[by_col]|Sorts a range or array.
SORTBY|array;by_array1;[sort_order1];...|Sorts a range or array based on the values in a corresponding range or array.
TAKE|array;rows;[columns]|Returns rows or columns from the start or end of an array.
TOCOL|array;[ignore];[scan_by_column]|Returns the array in a single column.
TOROW|array;[ignore];[scan_by_column]|Returns the array in a single row.
TRANSPOSE|array|Converts a vertical range of cells to a horizontal range, or vice versa.
UNIQUE|array;[by_col];[exactly_once]|Returns the unique values from a range or array.
VLOOKUP|lookup_value;table_array;col_index_num;[range_lookup]|Looks for a value in the leftmost column of a table, and then returns a value in the same row from a column you specify.
VSTACK|array1;[array2];...|Appends arrays vertically and in sequence.
WRAPCOLS|vector;wrap_count;[pad_with]|Wraps a row or column of values by columns after a specified number of elements.
WRAPROWS|vector;wrap_count;[pad_with]|Wraps a row or column of values by rows after a specified number of elements.
XLOOKUP|lookup_value;lookup_array;return_array;[if_not_found];[match_mode];[search_mode]|Searches a range or an array for a match and returns the corresponding item from a second range or array.
XMATCH|lookup_value;lookup_array;[match_mode];[search_mode]|Returns the relative position of an item in an array.`,
    Database: `
DAVERAGE|database;field;criteria|Averages the values in a column in a list or database that match conditions you specify.
DCOUNT|database;field;criteria|Counts the cells containing numbers in the field (column) of records in the database that match the conditions you specify.
DCOUNTA|database;field;criteria|Counts nonblank cells in the field (column) of records in the database that match the conditions you specify.
DGET|database;field;criteria|Extracts from a database a single record that matches the conditions you specify.
DMAX|database;field;criteria|Returns the largest number in the field (column) of records in the database that match the conditions you specify.
DMIN|database;field;criteria|Returns the smallest number in the field (column) of records in the database that match the conditions you specify.
DPRODUCT|database;field;criteria|Multiplies the values in the field (column) of records in the database that match the conditions you specify.
DSTDEV|database;field;criteria|Estimates the standard deviation based on a sample from selected database entries.
DSTDEVP|database;field;criteria|Calculates the standard deviation based on the entire population of selected database entries.
DSUM|database;field;criteria|Adds the numbers in the field (column) of records in the database that match the conditions you specify.
DVAR|database;field;criteria|Estimates variance based on a sample from selected database entries.
DVARP|database;field;criteria|Calculates variance based on the entire population of selected database entries.`,
    Text: `
ARRAYTOTEXT|array;[format]|Returns a text representation of an array.
ASC|text|Changes full-width (double-byte) characters to half-width (single-byte) characters.
BAHTTEXT|number|Converts a number to text (baht).
CHAR|number|Returns the character specified by the code number from the character set for your computer.
CLEAN|text|Removes all nonprintable characters from text.
CODE|text|Returns a numeric code for the first character in a text string, in the character set used by your computer.
CONCAT|text1;...|Concatenates a list or range of text strings.
CONCATENATE|text1;[text2];...|Joins several text strings into one text string.
DBCS|text|Changes half-width (single-byte) characters to full-width (double-byte) characters.
DOLLAR|number;[decimals]|Converts a number to text, using currency format.
EXACT|text1;text2|Checks whether two text strings are exactly the same, and returns TRUE or FALSE. EXACT is case-sensitive.
FIND|find_text;within_text;[start_num]|Returns the starting position of one text string within another text string. FIND is case-sensitive.
FINDB|find_text;within_text;[start_num]|Finds the starting position of one text string within another text string, counting bytes.
FIXED|number;[decimals];[no_commas]|Rounds a number to the specified number of decimals and returns the result as text with or without commas.
JIS|text|Changes half-width (single-byte) characters to full-width (double-byte) characters.
LEFT|text;[num_chars]|Returns the specified number of characters from the start of a text string.
LEFTB|text;[num_bytes]|Returns the specified number of bytes from the start of a text string.
LEN|text|Returns the number of characters in a text string.
LENB|text|Returns the number of bytes used to represent the characters in a text string.
LOWER|text|Converts all letters in a text string to lowercase.
MID|text;start_num;num_chars|Returns the characters from the middle of a text string, given a starting position and length.
MIDB|text;start_num;num_bytes|Returns bytes from the middle of a text string, given a starting position and length.
NUMBERVALUE|text;[decimal_separator];[group_separator]|Converts text to number in a locale-independent manner.
PHONETIC|reference|Extracts the phonetic (furigana) characters from a text string.
PROPER|text|Converts a text string to proper case; the first letter in each word in uppercase, and all other letters to lowercase.
REPLACE|old_text;start_num;num_chars;new_text|Replaces part of a text string with a different text string.
REPLACEB|old_text;start_num;num_bytes;new_text|Replaces part of a text string, counting bytes.
REPT|text;number_times|Repeats text a given number of times. Use REPT to fill a cell with a number of instances of a text string.
RIGHT|text;[num_chars]|Returns the specified number of characters from the end of a text string.
RIGHTB|text;[num_bytes]|Returns the specified number of bytes from the end of a text string.
SEARCH|find_text;within_text;[start_num]|Returns the number of the character at which a specific character or text string is first found, reading left to right (not case-sensitive).
SEARCHB|find_text;within_text;[start_num]|Finds one text value within another (counting bytes, not case-sensitive).
SUBSTITUTE|text;old_text;new_text;[instance_num]|Replaces existing text with new text in a text string.
T|value|Checks whether a value is text, and returns the text if it is, or returns double quotes (empty text) if it is not.
TEXT|value;format_text|Converts a value to text in a specific number format.
TEXTAFTER|text;delimiter;[instance_num];[match_mode];[match_end];[if_not_found]|Returns text that occurs after a given character or string.
TEXTBEFORE|text;delimiter;[instance_num];[match_mode];[match_end];[if_not_found]|Returns text that occurs before a given character or string.
TEXTJOIN|delimiter;ignore_empty;text1;...|Concatenates a list or range of text strings using a delimiter.
TEXTSPLIT|text;col_delimiter;[row_delimiter];[ignore_empty];[match_mode];[pad_with]|Splits text into rows or columns using delimiters.
TRIM|text|Removes all spaces from a text string except for single spaces between words.
UNICHAR|number|Returns the Unicode character referenced by the given numeric value.
UNICODE|text|Returns the number (code point) corresponding to the first character of the text.
UPPER|text|Converts a text string to all uppercase letters.
VALUE|text|Converts a text string that represents a number to a number.
VALUETOTEXT|value;[format]|Returns a text representation of a value.`,
    Logical: `
AND|logical1;[logical2];...|Checks whether all arguments are TRUE, and returns TRUE if all arguments are TRUE.
BYCOL|array;lambda|Applies a LAMBDA to each column and returns an array of the results.
BYROW|array;lambda|Applies a LAMBDA to each row and returns an array of the results.
FALSE||Returns the logical value FALSE.
IF|logical_test;[value_if_true];[value_if_false]|Checks whether a condition is met, and returns one value if TRUE, and another value if FALSE.
IFERROR|value;value_if_error|Returns value_if_error if expression is an error and the value of the expression itself otherwise.
IFNA|value;value_if_na|Returns the value you specify if the expression resolves to #N/A, otherwise returns the result of the expression.
IFS|logical_test1;value_if_true1;...|Checks whether one or more conditions are met and returns a value corresponding to the first TRUE condition.
LAMBDA|[parameter1];...;calculation|Creates a function value, which can be called within formulas.
LET|name1;name_value1;...;calculation|Assigns calculated values to names which allow storing intermediate calculation results.
MAKEARRAY|rows;columns;lambda|Returns an array of a specified row and column size, by applying a LAMBDA.
MAP|array1;...;lambda|Returns an array formed by mapping each value in the array(s) to a new value by applying a LAMBDA.
NOT|logical|Changes FALSE to TRUE, or TRUE to FALSE.
OR|logical1;[logical2];...|Checks whether any of the arguments are TRUE, and returns TRUE or FALSE. Returns FALSE only if all arguments are FALSE.
REDUCE|[initial_value];array;lambda|Reduces an array to an accumulated value by applying a LAMBDA to each value.
SCAN|[initial_value];array;lambda|Scans an array by applying a LAMBDA to each value and returns an array that has each intermediate value.
SWITCH|expression;value1;result1;...|Evaluates an expression against a list of values and returns the result corresponding to the first matching value.
TRUE||Returns the logical value TRUE.
XOR|logical1;[logical2];...|Returns a logical 'Exclusive Or' of all arguments.`,
    Information: `
CELL|info_type;[reference]|Returns information about the formatting, location, or contents of the first cell, according to the sheet's reading order, in a reference.
ERROR.TYPE|error_val|Returns a number corresponding to an error value.
INFO|type_text|Returns information about the current operating environment.
ISBLANK|value|Checks whether a reference is to an empty cell, and returns TRUE or FALSE.
ISERR|value|Checks whether a value is an error other than #N/A, and returns TRUE or FALSE.
ISERROR|value|Checks whether a value is an error, and returns TRUE or FALSE.
ISEVEN|number|Returns TRUE if the number is even.
ISFORMULA|reference|Checks whether a reference is to a cell containing a formula, and returns TRUE or FALSE.
ISLOGICAL|value|Checks whether a value is a logical value (TRUE or FALSE), and returns TRUE or FALSE.
ISNA|value|Checks whether a value is #N/A, and returns TRUE or FALSE.
ISNONTEXT|value|Checks whether a value is not text (blank cells are not text), and returns TRUE or FALSE.
ISNUMBER|value|Checks whether a value is a number, and returns TRUE or FALSE.
ISODD|number|Returns TRUE if the number is odd.
ISOMITTED|argument|Checks whether the value in a LAMBDA is missing and returns TRUE or FALSE.
ISREF|value|Checks whether a value is a reference, and returns TRUE or FALSE.
ISTEXT|value|Checks whether a value is text, and returns TRUE or FALSE.
N|value|Converts non-number value to a number, dates to serial numbers, TRUE to 1, anything else to 0 (zero).
NA||Returns the error value #N/A (value not available).
SHEET|[value]|Returns the sheet number of the referenced sheet.
SHEETS|[reference]|Returns the number of sheets in a reference.
TYPE|value|Returns an integer representing the data type of a value: number = 1; text = 2; logical value = 4; error value = 16; array = 64.`,
    Engineering: `
BESSELI|x;n|Returns the modified Bessel function In(x).
BESSELJ|x;n|Returns the Bessel function Jn(x).
BESSELK|x;n|Returns the modified Bessel function Kn(x).
BESSELY|x;n|Returns the Bessel function Yn(x).
BIN2DEC|number|Converts a binary number to decimal.
BIN2HEX|number;[places]|Converts a binary number to hexadecimal.
BIN2OCT|number;[places]|Converts a binary number to octal.
BITAND|number1;number2|Returns a bitwise 'And' of two numbers.
BITLSHIFT|number;shift_amount|Returns a number shifted left by shift_amount bits.
BITOR|number1;number2|Returns a bitwise 'Or' of two numbers.
BITRSHIFT|number;shift_amount|Returns a number shifted right by shift_amount bits.
BITXOR|number1;number2|Returns a bitwise 'Exclusive Or' of two numbers.
COMPLEX|real_num;i_num;[suffix]|Converts real and imaginary coefficients into a complex number.
CONVERT|number;from_unit;to_unit|Converts a number from one measurement system to another.
DEC2BIN|number;[places]|Converts a decimal number to binary.
DEC2HEX|number;[places]|Converts a decimal number to hexadecimal.
DEC2OCT|number;[places]|Converts a decimal number to octal.
DELTA|number1;[number2]|Tests whether two numbers are equal.
ERF|lower_limit;[upper_limit]|Returns the error function.
ERF.PRECISE|x|Returns the error function.
ERFC|x|Returns the complementary error function.
ERFC.PRECISE|x|Returns the complementary error function.
GESTEP|number;[step]|Tests whether a number is greater than a threshold value.
HEX2BIN|number;[places]|Converts a hexadecimal number to binary.
HEX2DEC|number|Converts a hexadecimal number to decimal.
HEX2OCT|number;[places]|Converts a hexadecimal number to octal.
IMABS|inumber|Returns the absolute value (modulus) of a complex number.
IMAGINARY|inumber|Returns the imaginary coefficient of a complex number.
IMARGUMENT|inumber|Returns the argument q, an angle expressed in radians.
IMCONJUGATE|inumber|Returns the complex conjugate of a complex number.
IMCOS|inumber|Returns the cosine of a complex number.
IMCOSH|inumber|Returns the hyperbolic cosine of a complex number.
IMCOT|inumber|Returns the cotangent of a complex number.
IMCSC|inumber|Returns the cosecant of a complex number.
IMCSCH|inumber|Returns the hyperbolic cosecant of a complex number.
IMDIV|inumber1;inumber2|Returns the quotient of two complex numbers.
IMEXP|inumber|Returns the exponential of a complex number.
IMLN|inumber|Returns the natural logarithm of a complex number.
IMLOG10|inumber|Returns the base-10 logarithm of a complex number.
IMLOG2|inumber|Returns the base-2 logarithm of a complex number.
IMPOWER|inumber;number|Returns a complex number raised to an integer power.
IMPRODUCT|inumber1;[inumber2];...|Returns the product of 1 to 255 complex numbers.
IMREAL|inumber|Returns the real coefficient of a complex number.
IMSEC|inumber|Returns the secant of a complex number.
IMSECH|inumber|Returns the hyperbolic secant of a complex number.
IMSIN|inumber|Returns the sine of a complex number.
IMSINH|inumber|Returns the hyperbolic sine of a complex number.
IMSQRT|inumber|Returns the square root of a complex number.
IMSUB|inumber1;inumber2|Returns the difference of two complex numbers.
IMSUM|inumber1;[inumber2];...|Returns the sum of complex numbers.
IMTAN|inumber|Returns the tangent of a complex number.
OCT2BIN|number;[places]|Converts an octal number to binary.
OCT2DEC|number|Converts an octal number to decimal.
OCT2HEX|number;[places]|Converts an octal number to hexadecimal.`,
    Cube: `
CUBEKPIMEMBER|connection;kpi_name;kpi_property;[caption]|Returns a key performance indicator (KPI) property and displays the KPI name in the cell.
CUBEMEMBER|connection;member_expression;[caption]|Returns a member or tuple from the cube.
CUBEMEMBERPROPERTY|connection;member_expression;property|Returns the value of a member property from the cube.
CUBERANKEDMEMBER|connection;set_expression;rank;[caption]|Returns the nth, or ranked, member in a set.
CUBESET|connection;set_expression;[caption];[sort_order];[sort_by]|Defines a calculated set of members or tuples by sending a set expression to the cube on the server.
CUBESETCOUNT|set|Returns the number of items in a set.
CUBEVALUE|connection;[member_expression1];...|Returns an aggregated value from the cube.`,
    Web: `
ENCODEURL|text|Returns a URL-encoded string.
FILTERXML|xml;xpath|Returns specific data from the XML content by using the specified XPath.
WEBSERVICE|url|Returns data from a web service.`,
  };
  const OLD = new Set('ACCRINT ACCRINTM AMORDEGRC AMORLINC COUPDAYBS COUPDAYS COUPDAYSNC COUPNCD COUPNUM COUPPCD CUMIPMT CUMPRINC DB DDB DISC DOLLARDE DOLLARFR DURATION EFFECT FV FVSCHEDULE INTRATE IPMT IRR ISPMT MDURATION MIRR NOMINAL NPER NPV ODDFPRICE ODDFYIELD ODDLPRICE ODDLYIELD PMT PPMT PRICE PRICEDISC PRICEMAT PV RATE RECEIVED SLN SYD TBILLEQ TBILLPRICE TBILLYIELD VDB XIRR XNPV YIELD YIELDDISC YIELDMAT DATE DATEVALUE DAY DAYS360 EDATE EOMONTH HOUR MINUTE MONTH NETWORKDAYS NOW SECOND TIME TIMEVALUE TODAY WEEKDAY WEEKNUM WORKDAY YEAR YEARFRAC ABS ACOS ACOSH ASIN ASINH ATAN ATAN2 ATANH CEILING COMBIN COS COSH DEGREES EVEN EXP FACT FACTDOUBLE FLOOR GCD INT LCM LN LOG LOG10 MDETERM MINVERSE MMULT MOD MROUND MULTINOMIAL ODD PI POWER PRODUCT QUOTIENT RADIANS RAND RANDBETWEEN ROMAN ROUND ROUNDDOWN ROUNDUP SERIESSUM SIGN SIN SINH SQRT SQRTPI SUBTOTAL SUM SUMIF SUMPRODUCT SUMSQ SUMX2MY2 SUMX2PY2 SUMXMY2 TAN TANH TRUNC AVEDEV AVERAGE AVERAGEA BETADIST BETAINV BINOMDIST CHIDIST CHIINV CHITEST CONFIDENCE CORREL COUNT COUNTA COUNTBLANK COUNTIF COVAR CRITBINOM DEVSQ EXPONDIST FDIST FINV FISHER FISHERINV FORECAST FREQUENCY FTEST GAMMADIST GAMMAINV GAMMALN GEOMEAN GROWTH HARMEAN HYPGEOMDIST INTERCEPT KURT LARGE LINEST LOGEST LOGINV LOGNORMDIST MAX MAXA MEDIAN MIN MINA MODE NEGBINOMDIST NORMDIST NORMINV NORMSDIST NORMSINV PEARSON PERCENTILE PERCENTRANK PERMUT POISSON PROB QUARTILE RANK RSQ SKEW SLOPE SMALL STANDARDIZE STDEV STDEVA STDEVP STDEVPA STEYX TDIST TINV TREND TRIMMEAN TTEST VAR VARA VARP VARPA WEIBULL ZTEST ADDRESS AREAS CHOOSE COLUMN COLUMNS GETPIVOTDATA HLOOKUP HYPERLINK INDEX INDIRECT LOOKUP MATCH OFFSET ROW ROWS RTD TRANSPOSE VLOOKUP DAVERAGE DCOUNT DCOUNTA DGET DMAX DMIN DPRODUCT DSTDEV DSTDEVP DSUM DVAR DVARP ASC BAHTTEXT CHAR CLEAN CODE CONCATENATE DOLLAR EXACT FIND FINDB FIXED JIS LEFT LEFTB LEN LENB LOWER MID MIDB PHONETIC PROPER REPLACE REPLACEB REPT RIGHT RIGHTB SEARCH SEARCHB SUBSTITUTE T TEXT TRIM UPPER VALUE AND FALSE IF NOT OR TRUE CELL ERROR.TYPE INFO ISBLANK ISERR ISERROR ISEVEN ISLOGICAL ISNA ISNONTEXT ISNUMBER ISODD ISREF ISTEXT N NA TYPE BESSELI BESSELJ BESSELK BESSELY BIN2DEC BIN2HEX BIN2OCT COMPLEX CONVERT DEC2BIN DEC2HEX DEC2OCT DELTA ERF ERFC GESTEP HEX2BIN HEX2DEC HEX2OCT IMABS IMAGINARY IMARGUMENT IMCONJUGATE IMCOS IMDIV IMEXP IMLN IMLOG10 IMLOG2 IMPOWER IMPRODUCT IMREAL IMSIN IMSQRT IMSUB IMSUM OCT2BIN OCT2DEC OCT2HEX'.split(' '));
  const map = new Map();
  for (const [cat, text] of Object.entries(DATA)) {
    for (const line of text.trim().split('\n')) {
      const [name, args, desc] = line.split('|');
      map.set(name, { name, cat, args: args ? args.split(';') : [], desc: desc || '', classic: OLD.has(name) });
    }
  }
  FI.get = (name) => map.get(String(name).toUpperCase()) || null;
  FI.all = () => Array.from(map.values());
  FI.categories = ['Most Recently Used', 'All', 'Financial', 'Date & Time', 'Math & Trig', 'Statistical', 'Lookup & Reference', 'Database', 'Text', 'Logical', 'Information', 'Engineering', 'Cube', 'Web'];
  FI.syntax = (f) => f.name + '(' + f.args.join(',') + ')';
  /* argument help shown in the Function Arguments dialog */
  const ARGHELP = {
    number: 'is a number, or a reference to a cell that contains a number.', number1: 'number1,number2,... are 1 to 255 numbers or arguments you want to process.',
    value1: 'value1,value2,... are 1 to 255 values or ranges.', logical1: 'logical1,logical2,... are 1 to 255 conditions you want to test that can be either TRUE or FALSE.',
    logical_test: 'is any value or expression that can be evaluated to TRUE or FALSE.', value_if_true: 'is the value that is returned if Logical_test is TRUE. If omitted, TRUE is returned.', value_if_false: 'is the value that is returned if Logical_test is FALSE. If omitted, FALSE is returned.',
    lookup_value: 'is the value to be found in the first column (or row) of the table, and can be a value, a reference, or a text string.', table_array: 'is a table of text, numbers, or logical values, in which data is retrieved. Table_array can be a reference to a range or a range name.',
    col_index_num: 'is the column number in table_array from which the matching value should be returned. The first column of values in the table is column 1.', row_index_num: 'is the row number in table_array from which the matching value should be returned.',
    range_lookup: 'is a logical value: to find the closest match in the first column (sorted in ascending order) = TRUE or omitted; find an exact match = FALSE.',
    range: 'is the range of cells you want evaluated.', criteria: 'is the condition or criteria in the form of a number, expression, or text that defines which cells will be added.', sum_range: 'are the actual cells to sum. If omitted, the cells in range are used.',
    text: 'is the text string you want processed.', num_chars: 'specifies how many characters you want.', rate: 'is the interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR.',
    nper: 'is the total number of payment periods in an investment.', pv: 'is the present value: the total amount that a series of future payments is worth now.', fv: 'is the future value, or a cash balance you want to attain after the last payment is made.', type: 'is a logical value: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted.',
    num_digits: 'is the number of digits to which you want to round. Negative rounds to the left of the decimal point; zero to the nearest integer.', serial_number: 'is a number in the date-time code used by Ledger.',
    array: 'is a range of cells or an array constant.', k: 'is the position (from the largest or smallest) in the array or cell range of data to return.', reference: 'is the cell or range of cells.',
    database: 'is the range of cells that makes up the list or database. A database is a list of related data.', field: 'is either the label of the column in double quotation marks or a number that represents the column\'s position in the list.', criteria_range1: 'is the range of cells you want evaluated for the particular condition.',
  };
  FI.argHelp = (name) => ARGHELP[String(name).replace(/[[\]]|\.\.\./g, '')] || '';
})(typeof window !== 'undefined' ? window : globalThis);
