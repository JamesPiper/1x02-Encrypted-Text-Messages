/////////////////////////////////////////////////////////////////////////////////////
// Project     : 1x02 Encrypted Text Messages
// Author      : James Piper, james@jamespiper.com
// Date        : 2017.01.22
// File        : _1x01_Enter_Plaintext.c
// Description : Get and store plaintext message.
// IDE         : Code::Blocks 16.01
// Compiler    : GCC
// Language    : C (Compiling to ISO 11)
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
// Function prototypes.
/////////////////////////////////////////////////////////////////////////////////////
int SetPlaintext(char* text);

/////////////////////////////////////////////////////////////////////////////////////
// Main function
/////////////////////////////////////////////////////////////////////////////////////
void _1x01_Enter_Plaintext() {

    /////////////////////////////////////////////////////////////////////////////////////
    // PROCEESS
    // 1. The objective is to get user-entered plaintext.
    //    In this implementation, a user will type data into a terminal window.
    //    In other implementations, the data can be passed by function call.
    // 2. The specifications limit the text message to 140 A to Z characters.
    //    The data can therefore be easily stored in an array of char.
    // 3. Parse text to remove non-alpha characters.
    //    From : It's a message.
    //    To   : Itsamessage
    // 4. Transform text to uppercase.
    //    To   : ITSAMESSAGE
    // 5. Check the message is no more than 140 characters.
    //
    // RETURN
    // 1. Formatted plaintext to a max of 140 chars.
    // 2. Indicate length of free space to reach max of 140.
    //    Example: len(plaintext) = 100; free = 40;
    // 3. Otherwise return negative int for invalid plaintext.
    //
    // definition: int SetPlaintext(char* text);
    //
    // OTHER
    // Allow for adjustment to the message - reduce, increase, edit...
    // Using this project to do that is not easy, better to do
    // with an implementation in a GUI environment (Windows, cell phone app).
    //
    /////////////////////////////////////////////////////////////////////////////////////

	printf("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n");
	printf("Plaintext Input\n");
	printf("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n");
	printf("\n");
	printf("\n");
	printf("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n");
    system("pause");

}

int SetPlaintext(char* text) {

    char* pText = (char*) malloc(sizeof(text));
    return 0;

}

