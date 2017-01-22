/////////////////////////////////////////////////////////////////////////////////////
// Project     : 1x02 Encrypted Text Messages
// Author      : James Piper, james@jamespiper.com
// Date        : 2017.01.22
// File        : 1x02 Encrypted Text Messages (Lib).c
// Description : Library of general purpose functions.
// IDE         : Code::Blocks 16.01
// Compiler    : GCC
// Language    : C (Compiling to ISO 11.)
/////////////////////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////////////////////
// Macros
/////////////////////////////////////////////////////////////////////////////////////
//#define DEBUG

/////////////////////////////////////////////////////////////////////////////////////
// Include files
/////////////////////////////////////////////////////////////////////////////////////
#include "1x02 Encrypted Text Messages.h"
/////////////////////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////////////////////
// Library functions.
/////////////////////////////////////////////////////////////////////////////////////
Boolean FileExists(const char* filename) {

	// There seems to be several ways to do this, but using fopen seems the easiest.
	// One method uses access() yet I can't find any reference on this.
	// Another uses stat() function in <sys/stat.h>
	// Using const in argument so this function doesn't change the filename.

	FILE* pFile;
	pFile = fopen(filename, "r");
	if (pFile != NULL) {
		fclose(pFile);
		return True;
	} else
		return False;

}

int StringCompare(const char* string1, const char* string2) {

	/////////////////////////////////////////////////////////////////////////////////////
	// Compare string1 to string2 without regard to case.
	// Thus ABC is equal to abc or AbC and the other variations.
	// Also, apple comes before apples.
	//
	// Returns
	// 0 if string1 is the same as string2
	// < 0 if string1 comes before string2
	// > 0 if string1 comes after string2
	//
	// Risks or Concerns
	// Reading past array.
	// Doesn't deal with wide chars.
	// Leading or trailing whitespace - should be ignored.
	//
	/////////////////////////////////////////////////////////////////////////////////////
	//
	/////////////////////////////////////////////////////////////////////////////////////

	// To avoid reading past an array.
	size_t Length = strlen(string1);
	size_t LengthA = strlen(string1);
	size_t LengthB = strlen(string2);
	if (LengthB < LengthA)
		Length = LengthB;

	char A, B;
	int result = 0;

	for (int i = 0; i < Length; i++) {
		A = tolower(string1[i]);
		B = tolower(string2[i]);
		// Want to keep result from previous compares.
		if (A == B)
			result = abs(result) * result;
		else if (A <= B)
			result = -1;
		else if (A >= B)
			result = 1;
	}

	// Handle cases with different length strings.
	if (LengthA < LengthB) {
		// Case of 'app' v 'apps'
		if (result == 0)
			result = -1;
	} else if (LengthA > LengthB) {
		// Case of 'apps' v 'app'
		if (result == 0)
			result = 1;
	}

	return result;

}

char* TrimWhitespace(char* string) {

	/////////////////////////////////////////////////////////////////////////////////////
	// Take a string and remove leading and trailing spaces.
	//
	// The arguement string is changed if needed.
	//
	// Risks
	// Overflow: reading past the end of the string array or before it.
	// String doesn't end with '\0'.
	// String is too long. What is the limit?
	// String is null.
	// String is only whitespace.
	//
	/////////////////////////////////////////////////////////////////////////////////////
	//
	// From: " some (gap) text " + '\0'
	//        01234567890123456     7
	// To:   "some (gap) text"   + '\0'
	//        012345678901234       5   67
	// End:                   6
	//
	/////////////////////////////////////////////////////////////////////////////////////

	Boolean EndReached = False;
	int From = 0;
	int To = 0;
	size_t End = strlen(string) - 1;
	Boolean InText = False;

	// Deal with leading spaces.
	while (EndReached != True) {
		if (From < End) {
			if (isspace(string[From]) != 0) {
				if (InText)
					string[To++] = string[From++];
				else
					From++;
			} else {
				InText = True;
				string[To++] = string[From++];
			}
		} else {
			string[To++] = string[From++];
			EndReached = True;
		}
	}

	/////////////////////////////////////////////////////////////////////////////////////
	//printf("To:   %d \n", To);
	//printf("End:  %d \n", End);
	//printf("From: %d \n", From);
	/////////////////////////////////////////////////////////////////////////////////////

	// Backward search for trailing spaces.
	EndReached  = False;
	while ((EndReached != True) && (To > 0)) {
		 if (isspace(string[To - 1]) != 0)
			To--;
		 else {
			EndReached = True;
		 }
	}

	string[To] = '\0';
	return string;
}

void GetUserInputs(char* input, int max_length) {

	/////////////////////////////////////////////////////////////////////////////////////
	// Take user typed-in data from terminal and stores the string of inputs in
	// a string array, input, of max_length.
	//
	// Using scanf("%s", Inputs) doesn't work when there's whitespace.
	// It will only put first chunk of text into Inputs.
	//
	// There's many things said about overflows and how it can crash a program or worse.
	// Certainly if you try to store text with n chars into a char array defined with
	// m chars and m is less than n, you'll overwrite memory beyond the array or
	// you'll get a segmentation fault.
	//
	// Risks
	// 1. User inputs more text than max_length.
	//    Circumvent by ingoring text entereded after max_length and not storing data
	//    in array past max_length.
	//
	// 2. String doesn't end with '\0'.
	//
	// 3. scanf("%s", Inputs) on first entry.
	//    The computer will take a chunk of text until the whitespace is encountered
	//    and put it in Inputs.
	//    Possible overflow in writing to Inputs array if this chunk is longer.
	//    Use scanf("%ns", Inputs) to limit chars put into Inputs to n chars.
	//    Coding to vary this specifier based on max_length.
	//
	/////////////////////////////////////////////////////////////////////////////////////
	//
	// Example
	//
	// max_length: 10
	// Note: String defined in caller as str[10]
	//       10 = 9 chars + 1 of '\0'
	//
	// Enter text: 'here is some text' <enter>
	//              01234567890123456
	//
	// 1. input = "here" + \0
	// 2. i = 5
	// 3. get additional char until enter key hit
	// 4. add if not before max_length
	// 5. input = "here is s" + '\0'
	//             012345678     9
	//
	/////////////////////////////////////////////////////////////////////////////////////

	// First chunk of non-whitespace user input.
	// Using fixed length to simplify code.
	// Use more temp memory for less cycles.
	// If user enters more than 1024 chars, the code will likely cause problems.
	// Low risk of overflow.
	// Issue if piping in file as input.
	char UserInput[1024];
	scanf("%1023s", UserInput);
	strncpy(input, UserInput, max_length - 1);
	//*(input + max_length - 1) = '\0';

	// Initial chunk of text--cut off from first whitespace.
	size_t i = strlen(input);
	char c;

	do {
		// Get additional text one character at a time.
        scanf("%c", &c);

		// Keep adding text as long before max_length.
        if (i < max_length)
            input[i++] = c;

		// Need to continue looping until '\n' is reached
		// because additional text input will be processed.
		// It would be good to stop the scan process but I
		// don't know how to do that.
	} while (c != '\n');

	// Terminate string array.
	if (i < max_length)
		input[i - 1] = '\0';
	else
		input[max_length - 1] = '\0';

}

