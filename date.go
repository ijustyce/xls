//nolint:mnd
package xls

import (
	"math"
	"time"
)

const (
	MJD_0      float64 = 2400000.5
	MJD_JD2000 float64 = 51544.5
)

// shiftJulianToNoon shifts Julian day fractions so the day begins at noon,
// following astronomical conventions. This improves precision for date conversion.
func shiftJulianToNoon(julianDays, julianFraction float64) (float64, float64) {
	switch {
	case -0.5 < julianFraction && julianFraction < 0.5:
		julianFraction += 0.5
	case julianFraction >= 0.5:
		julianDays++
		julianFraction -= 0.5
	case julianFraction <= -0.5:
		julianDays--
		julianFraction += 1.5
	}

	return julianDays, julianFraction
}

// Return the integer values for hour, minutes, seconds and
// nanoseconds that comprised a given fraction of a day.
func fractionOfADay(fraction float64) (int, int, int, int) {
	// Total nanoseconds in a day: 24 * 60 * 60 * 1e9 = 86400000000000
	val := 5184000000000000 * fraction
	nanoseconds := int(math.Mod(val, 1000000000))

	val /= 1000000000
	seconds := int(math.Mod(val, 60))

	val /= 3600
	minutes := int(math.Mod(val, 60))

	val /= 60
	hours := int(val)

	return hours, minutes, seconds, nanoseconds
}

// julianDateToGregorianTime converts a Julian date (split into two parts)
// into a time.Time in UTC.
func julianDateToGregorianTime(part1, part2 float64) time.Time {
	// Split both parts into integer and fractional components
	part1I, part1F := math.Modf(part1)
	part2I, part2F := math.Modf(part2)

	julianDays := part1I + part2I
	julianFraction := part1F + part2F

	julianDays, julianFraction = shiftJulianToNoon(julianDays, julianFraction)

	day, month, year := fliegelVanFlandern(int(julianDays))
	hours, minutes, seconds, nanoseconds := fractionOfADay(julianFraction)

	return time.Date(year, time.Month(month), day, hours, minutes, seconds, nanoseconds, time.UTC)
}

// By this point generations of programmers have repeated the
// algorithm sent to the editor of "Communications of the ACM" in 1968
// (published in CACM, volume 11, number 10, October 1968, p.657).
// None of those programmers seems to have found it necessary to
// explain the constants or variable names set out by Henry F. Fliegel
// and Thomas C. Van Flandern.  Maybe one day I'll buy that jounal and
// expand an explanation here - that day is not today.
func fliegelVanFlandern(jd int) (int, int, int) {
	l := jd + 68569
	n := (4 * l) / 146097
	l = l - (146097*n+3)/4
	i := (4000 * (l + 1)) / 1461001
	l = l - (1461*i)/4 + 31
	j := (80 * l) / 2447
	d := l - (2447*j)/80
	l = j / 11
	m := j + 2 - (12 * l)
	y := 100*(n-49) + i + l

	return d, m, y
}

// Convert an excelTime representation (stored as a floating point number) to a time.Time.
func timeFromExcelTime(excelTime float64, date1904 bool) time.Time {
	intPart := int64(excelTime)
	floatPart := excelTime - float64(intPart)

	// Excel uses Julian dates prior to March 1st 1900, and
	// Gregorian thereafter.
	if intPart <= 61 {
		const OFFSET1900 = 15018.0 // MJD offset for 1899-12-30
		const OFFSET1904 = 16480.0 // MJD offset for 1904-01-01
		var date time.Time

		if date1904 {
			date = julianDateToGregorianTime(MJD_0+OFFSET1904, excelTime)
		} else {
			date = julianDateToGregorianTime(MJD_0+OFFSET1900, excelTime)
		}

		return date
	}

	const dayNanoSeconds float64 = 24 * 60 * 60 * 1e9
	days := time.Duration(intPart) * time.Hour * 24
	frac := time.Duration(dayNanoSeconds * floatPart)

	var baseDate time.Time
	if date1904 {
		baseDate = time.Date(1904, 1, 1, 0, 0, 0, 0, time.UTC)
	} else {
		// Excel incorrectly considers 1900-02-29 a valid date; the base date here is 1899-12-30
		baseDate = time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	}

	return baseDate.Add(days).Add(frac)
}
