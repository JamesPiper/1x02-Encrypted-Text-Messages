/////////////////////////////////////////////////////////////////////////////////////
// Project     : 1x02 Encrypted Text Messages
// Author      : James Piper, james@jamespiper.com
// Date        : 2017.01.22
// File        : _0x00_MainMenu.c
// Description : Terminal style main menu for user.
//             : Starting point for the user.
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
// Main function
/////////////////////////////////////////////////////////////////////////////////////
void _0x00_MainMenu() {

	char Choice;
    char* Message;
    Message = (char*) malloc(sizeof(MAX_LENGTH_OF_MESSAGE + 1));

	do
	{
		printf("******************************************************************************\n");
		printf("*                                                                            *\n");
		printf("*   Encrypted Text Messages                                                  *\n");
		printf("*   Main Menu                                                                *\n");
		printf("*                                                                            *\n");
		printf("*   Type Character + Enter                                                   *\n");
		printf("*                                                                            *\n");
		printf("*   P - Enter Plaintext                                                      *\n");
		printf("*   C - Enter Cyphertext                                                     *\n");
		printf("*   K - Enter Key                                                            *\n");
		printf("*                                                                            *\n");
		printf("*   E - Encrypt Message                                                      *\n");
		printf("*   D - Decrypt Message                                                      *\n");
		printf("*                                                                            *\n");
		printf("*   G - Generate Set of Keys                                                 *\n");
		printf("*   V - View Security Protocol                                               *\n");
		printf("*                                                                            *\n");
		printf("*   X - Exit the program.                                                    *\n");
		printf("*                                                                            *\n");
		printf("******************************************************************************\n");

		printf("\n");
		printf("Enter choice: ");

		// Input user action.
        char Inputs[MAX_INPUT_CHARS];
		GetUserInputs(Inputs, CHOICE_LENGTH);
		Choice = tolower(Inputs[0]);
		printf("\n");

        char UnitTest = 'n';
		#ifdef DEBUG
		if (Choice != 'x') {
            printf("Run unit test (Y/N)? ");
            strcpy(Inputs, " ");
            GetUserInputs(Inputs, CHOICE_LENGTH);
            UnitTest = tolower(Inputs[0]);
            printf("Unit test %c\n", UnitTest);
            printf("\n");
           if ((UnitTest != 'n') && (UnitTest != 'y'))
                UnitTest = 'n';
        }
        #endif // DEBUG

        if (UnitTest == 'n') {
            if (Choice == 'p')
                _1x01_Enter_Plaintext();
            else if (Choice == 'c')
                _1x02_Enter_Cyphertext();
            else if (Choice == 'k')
                _1x03_Enter_Key();
            else if (Choice == 'e')
                _1x04_Encrypt_Message();
            else if (Choice == 'd')
                _1x05_Decrypt_Message();
            else if (Choice == 'g')
                _1x06_Generate_Key();
            else if (Choice == 'v')
                Choice = 'v';
            else if (Choice != 'x')
                printf("*** Select a choice from those listed. ****\n\n");
        } else {
            if (Choice == 'p')
                Choice = 'p';
            else if (Choice == 'c')
                Choice = 'c';
            else if (Choice == 'k')
                Choice = 'k';
            else if (Choice == 'e')
                Choice = 'e';
            else if (Choice == 'd')
                Choice = 'd';
            else if (Choice == 'v')
                Choice = 'v';
            else
                printf("*** Select a choice from those listed. ****\n\n");
        }
	} while (Choice != 'x');
}
