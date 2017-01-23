#include "stdafx.h"
#define _CRT_SECURE_NO_WARNINGS

/////////////////////////////////////////////////////////////////////////////////
// Help
/////////////////////////////////////////////////////////////////////////////////

// Standard I/O library.
#include <cstdio>

// The standard library includes the system function.
#include <cstdlib>

int DisplayHelp()
{

	printf("Help.\n");

	// File handle.
	FILE * pFile;

	// Source file with help info.
	char * Filename = "Encrypted Message Intro Page.txt";

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

	// Variables for reading the file
	const short MaxFileLine = 1280;
	char sLine[MaxFileLine];

	// Variables for word wrapping.
	const short LineWidth = 45;
	short Start = 0;
	short End = 0;
	short PrevSpace = 0;
	short CurrSpace = 0;
	short i = 0; 
	short j = 0;
	short WrapCount = 0;

	// Loop through the file with text on help to display.
	while (!feof(pFile))
	{
		// Read one line of text from file.
		fgets(sLine, MaxFileLine, pFile);

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
	}

	// Close.
	short iReturn;
	iReturn = fclose(pFile);

	// Pause.
	system("pause");

	// Return success.
	return 0;
}