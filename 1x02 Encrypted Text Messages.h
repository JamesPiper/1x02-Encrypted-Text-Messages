/////////////////////////////////////////////////////////////////////////////////////
// Project     : 1x02 Encrypted Text Messages
// Author      : James Piper, james@jamespiper.com
// Date        : 2017.01.22
// File        : 1x02 Encrypted Text Messages.h
// Description : Main header file.
// IDE         : Code::Blocks 16.01
// Compiler    : GCC
// Language    : C (Compiling to ISO 11.)
/////////////////////////////////////////////////////////////////////////////////////
//
// https://en.wikipedia.org/wiki/C_preprocessor
// http://www.cprogramming.com/tutorial/cpreprocessor.html
//
/////////////////////////////////////////////////////////////////////////////////////

//#pragma once

#ifndef MAIN_HEADER_FILE
#define MAIN_HEADER_FILE

/////////////////////////////////////////////////////////////////////////////////////
// Include files
/////////////////////////////////////////////////////////////////////////////////////
#include <stdlib.h>
#include <stdio.h>
#include <ctype.h>
#include <string.h>
#include <math.h>
/////////////////////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////////////////////
// Macros
/////////////////////////////////////////////////////////////////////////////////////
#define MAX_INPUT_CHARS 255
#define CHOICE_LENGTH 2

/////////////////////////////////////////////////////////////////////////////////////
// Common typedefs
/////////////////////////////////////////////////////////////////////////////////////
typedef enum Boolean { False, True } Boolean;

/////////////////////////////////////////////////////////////////////////////////////
// Function prototypes.
/////////////////////////////////////////////////////////////////////////////////////
void _0x00_MainMenu();
void _1x00_Encrypt_Message();
void _1x01_Decrypt_Message();
void _1x02_Enter_Key();
void _1x03_Generate_Key();

// Library functions.
Boolean FileExists(const char* filename);
int StringCompare(const char* string1, const char* string2);
char* TrimWhitespace(char* string);
void GetUserInputs(char* inputs, int);


#endif // HEADER_FILE

