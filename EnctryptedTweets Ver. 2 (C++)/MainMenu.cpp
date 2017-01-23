#include "stdafx.h"
#include "EnctryptedTweets Ver. 2.h"
#define _CRT_SECURE_NO_WARNINGS

/////////////////////////////////////////////////////////////////////////////////
// Main Menu
/////////////////////////////////////////////////////////////////////////////////

// Standard I/O library.
#include <cstdio>


int MainMenu()
{
	// Input variable.
	char InputAll[80];
	char Action = 'Q';

	do
	{

		printf("******************************************\n");
		printf("*        Encrypt Message Program         *\n");
		printf("*               Main Menu                *\n");
		printf("*                                        *\n");
		printf("*  Type Character                        *\n");
		printf("*                                        *\n");
		printf("*   C - Enter random card data.          *\n");
		printf("*   L - List avaible card data.          *\n");
		printf("*   E - Encipher message.                *\n");
		printf("*   D - Decipher message.                *\n");
		printf("*   S - Save message details.            *\n");
		printf("*   ? - Help.                            *\n");
		printf("*   Q - Exit the program.                *\n");
		printf("*                                        *\n");
		printf("******************************************\n");

		printf("\n");
		printf("\n");
		printf("Enter action: ");

		// Input user action.
		// Because of buffering, getchar or getc processes hold more than one character.
		// Use scanf because less chance of problems.

		scanf("%s", &InputAll);
		// printf("\n%s", &InputAll);

		// Simple parse input.
		Action = InputAll[0];

		if (Action == 'C' || Action == 'c') 
		{
			// Adding cards.
			AddCards();
		}
		else if  (Action == 'E' || Action == 'e') 
		{
			// Encipher msg.
			EncipherMsg();
		}
		else if  (Action == 'L' || Action == 'l') 
		{
			// List available cards for use as key.
			ListCards();
		}
		else if  (Action == 'D' || Action == 'd') 
		{
			// Decipher msg.
			DecipherMsg();
		}
		else if  (Action == '?') 
		{
			// Help page.
			DisplayHelp();
		}
		else 
		{
			// Other.
			printf("Invalid input.\n");
		}

		printf("\n");

	} while (Action != 'Q' && Action != 'q');
	
	return 0;
}