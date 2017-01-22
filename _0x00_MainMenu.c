/////////////////////////////////////////////////////////////////////////////////////
// Project     : 1x02 Encrypted Text Messages
// Author      : James Piper, james@jamespiper.com
// Date        : 2017.01.22
// File        : _0x00_MainMenu.c
// Description : Terminal style main menu for user.
//             : Starting point for the user.
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
// Main function
/////////////////////////////////////////////////////////////////////////////////////
void _0x00_MainMenu() {

	char Choice;

	do
	{
		printf("******************************************************************************\n");
		printf("*                                                                            *\n");
		printf("*   Encrypted Text Messages                                                  *\n");
		printf("*   Main Menu                                                                *\n");
		printf("*                                                                            *\n");
		printf("*   Type Character + Enter                                                   *\n");
		printf("*                                                                            *\n");
		printf("*   A - Encrypt Message                                                      *\n");
		printf("*   B - Decrypt Message                                                      *\n");
		printf("*   C - Enter Key                                                            *\n");
		printf("*   D - Generate Key                                                         *\n");
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

		if (Choice == 'a') {
            if (UnitTest == 'n')
                Choice = 'a';
            else
                Choice = 'a';
		} else if (Choice == 'b') {
            if (UnitTest == 'n')
                Choice = 'b';
            else
                Choice = 'b';
		} else if (Choice == 'c')
                Choice = 'c';
		else if (Choice == 'd')
            Choice = 'd';
		//else if (Choice == 'e')
        //    Choice = 'e';
        else if (Choice != 'x')
			printf("*** Select a choice from those listed. ****\n\n");

	} while (Choice != 'x');

}
