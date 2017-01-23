/////////////////////////////////////////////////////////////////////////////////////
// Project     : 1x02 Encrypted Text Messages
// Author      : James Piper, james@jamespiper.com
// Date        : 2017.01.22
// File        : _1x04_Encrypt_Message.c
// Description : Encrypt a user-entered text message.
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
void _1x04_Encrypt_Message() {

    /////////////////////////////////////////////////////////////////////////////////////
    // PROCEESS
    // 1. Check if plaintext exists, don't continue if it doesn't.
    // 2. Check if cypher key exists, don't continue if it doesn't.
    // 3. Check if key is at least as long as the message.
    //    Options on error.
    //    (a) get longer key
    //    (b) limit cypertext to length of key
    //    (c) do not encrypt
    // 4. Create cypertext from plaintext and key.
    // 5. Test that cypertext is different from plaintext.
    // 6. Output cyphertext.
    // 7. Commit reduction in key set for cypher key characters used.
    //    Risk key set is not reduced creating out of sync sets.
    //
    // OPTIONS
    // 1. Display cyphertext in 5 char + 1 space.
    //    Off : OWJDIHAOAIXUYMMW
    //    On  : OWJDI HAOAI XUYMM W
    //
    // 2. Add one character padding.
    //    Off : OWJDI HAOAI XUYMM W
    //    On  : OLWYJ WDSIB HKANO NATIV XXUWY ZMTMY W
    //
    /////////////////////////////////////////////////////////////////////////////////////

	printf("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n");
	printf("Encrypt Message\n");
	printf("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n");
	printf("\n");
	printf("\n");
	printf("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n");
    system("pause");

}