package main

import "testing"

type testcase struct {
	cellRange string
	expectedColumnStart int
	expectedColumnEnd int
	expectedRowStart int
	expectedRowEnd int
}

var tests = []testcase {
	{ "A1:B2", 1,2,1,2},
	{ "B2:A1", 1,2,1,2},
	{ "B1:A3", 1,2,1,3},
	{ "AB12:C3", 3,28,3,12},
}

func TestParseRange(t *testing.T) {
	for _, tc := range tests {
		columnStart, columnEnd, rowStart, rowEnd := parseRange(tc.cellRange)
		t.Log("Testing cell range", tc.cellRange)
		checkValue(tc.expectedColumnStart, columnStart, "columnStart", t)
		checkValue(tc.expectedColumnEnd, columnEnd, "columnEnd", t)
		checkValue(tc.expectedRowStart, rowStart, "rowStart", t)
		checkValue(tc.expectedRowEnd, rowEnd, "rowEnd", t)
	}
}

func checkValue(expected int, actual int, valueName string, t *testing.T) {
	if actual != expected {
		t.Error(
		"For property", valueName,
		"expected", expected,
		"got", actual)
	}
}