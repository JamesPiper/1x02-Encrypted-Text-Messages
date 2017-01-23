/////////////////////////////////////////////////////////////////////////////////////
// Project     : 1x02 Encrypted Text Messages
// Author      : James Piper, james@jamespiper.com
// Date        : 2017.01.22
// File        : _1x02_Enter_Cyphertext.c
// Description : Get and store cyphertext message.
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
int SetCyphertext(char* text);

/////////////////////////////////////////////////////////////////////////////////////
// Main function
/////////////////////////////////////////////////////////////////////////////////////
void _1x02_Enter_Cyphertext() {

    /////////////////////////////////////////////////////////////////////////////////////
    // PROCEESS
    // 1. The objective is to get user-entered cyphertext.
    //    In this implementation, a user will type data into a terminal window.
    //    In other implementations, the data can be passed by function call.
    // 2. The cyphertext should be a series of alpha characters.
    //    There may be spaces.
    //    There may be padding--every other character is ignored.
    // 3. Transform text to uppercase.
    //    Remove any spaces.
    //    From : itgbj vyshz vysjw ykfgs esmxu nvnt
    //    To   : IGJYHVSWKGEMUVT
    //
    // RETURN
    // 1. Formatted cyphertext to a max of 140 chars.
    // 2. Length of message.
    // 3. Otherwise return negative int for invalid cyphertext.
    //
    // definition: int SetCyphertext(char* text);
    //
    /////////////////////////////////////////////////////////////////////////////////////

	printf("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n");
	printf("Cyphertext Input\n");
	printf("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n");
	printf("\n");
	printf("\n");
	printf("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n");
    system("pause");

}

int SetCyphertext(char* text) {

    char* pText = (char*) malloc(sizeof(text));
    return strlen(pText);

}

