/////////////////////////////////////////////////////////////////////////////////////
// Project     : 1x02 Encrypted Text Messages
// Author      : James Piper, james@jamespiper.com
// Date        : 2017.01.22
// File        : _1x03_Enter_Key.c
// Description : Enter key.
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

/////////////////////////////////////////////////////////////////////////////////////
// Main function
/////////////////////////////////////////////////////////////////////////////////////
void _1x03_Enter_Key() {

    /////////////////////////////////////////////////////////////////////////////////////
    // PROCEESS
    // 1. Get the cypher key to use for encrypting or decrypting.
    //    Two possible sources.
    //    (a) entered by user
    //    (b) stored in a keyset file
    //
    // 2. Check if key is at least as long as the message.
    //    May be more appropriate in encrypt or decrypt functions.
    //    This requires setting the plaintext or cyphertext first.
    //
    //    Options on error.
    //    (a) get longer key
    //    (b) limit plaintext or cyphertext to length of key
    //    (c) do not encrypt
    //
    // 3. Test randomness of key.
    //    May be more appropriate in generating the key set.
    //    Difficult to determine (a) the key is short (b) what is random.
    //    Advanced algorithm required re: letter frequency, sequencing.
    //    This could or should be part of entering key or in creating set of keys.
    //
    /////////////////////////////////////////////////////////////////////////////////////

	printf("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n");
	printf("Cypher Key Input\n");
	printf("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n");
	printf("\n");
	printf("\n");
	printf("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n");
    system("pause");

}