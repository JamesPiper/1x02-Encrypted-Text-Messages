#include "stdafx.h"
#define _CRT_SECURE_NO_WARNINGS

/////////////////////////////////////////////////////////////////////////////////
// List of Available Cards
/////////////////////////////////////////////////////////////////////////////////

// Standard I/O library.
#include <cstdio>

// The standard library includes the system function.
#include <cstdlib>

int ListCards()
{
	/////////////////////////////////////////////////////////////////////////////
	// Open file.

	// File handle.
	FILE * pFile;

	// Source file with help info.
	char * Filename = "AvailableCardValues.dat";

	// Open the file for reading.
	pFile = fopen(Filename, "r");

	// Test if failure.
	if(pFile == NULL)
	{
		// Err msg.
		perror("Unable to open the help file ");
		// Failure return.
		return -1;
	}

	/////////////////////////////////////////////////////////////////////////////
	// Header.
	printf("\n\nList of Available Cards\n\n");

	/////////////////////////////////////////////////////////////////////////////
	// Read card data from the file.

	const unsigned short MaxFileLine = 65535;
	char sLine[MaxFileLine];
	short i = 0; 
	char c;

	// Loop through the file.
	while (!feof(pFile))
	{
		c = fgetc(pFile);
		if (c != '\n')
		{
			sLine[i++] = c;
		}
	}

	// End array with newline & null char.
	sLine[i++] = '\n';
	sLine[i++] = '\0';

	/////////////////////////////////////////////////////////////////////////////
	// Display file contents with word wrapping.

	// Variables for word wrapping.
	const short LineWidth = 75;
	short Start = 0;
	short End = 0;
	short PrevSpace = 0;
	short CurrSpace = 0;
	short j = 0;
	short WrapCount = 0;

	// Reset counters.
	Start = 0;
	End = 0;
	PrevSpace = 0;
	CurrSpace = 0;
	WrapCount = 0;

	// Loop through characters in the line.
	for (i = 0; sLine[i] ; i++)

		// Look for space to find spot for word wrapping.
		if (sLine[i] == ' ') 
		{
			// Move to next space.
			PrevSpace = CurrSpace;
			CurrSpace = i;

			// Check if at end of line.
			if (CurrSpace > (LineWidth * (WrapCount + 1))) 
			{
				// Wrap line before or after current space?
				if ((CurrSpace - LineWidth) > (PrevSpace - LineWidth))
					End = PrevSpace;
				else
					End = CurrSpace;

				// Output the first line of text.
				for (j = Start; j <= End; j++)
					putchar(sLine[j]);
				// Add new line.
				putchar('\n');
				// Move start.
				Start = End + 1;
				// Increase wrap count.
				WrapCount++;
			}
		} else if (sLine[i] == '\n') 
		{
			for (j = Start; sLine[j]; j++)
				putchar(sLine[j]);
		}

	/////////////////////////////////////////////////////////////////////////////
	// Close.
	short iReturn;
	iReturn = fclose(pFile);

	// Pause so user can see the list before main menu is viewed.
	system("pause");

	// Return success.
	return 0;
}