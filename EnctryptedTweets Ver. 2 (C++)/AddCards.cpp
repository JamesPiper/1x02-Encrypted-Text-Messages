#include "stdafx.h"

/////////////////////////////////////////////////////////////////////////////////
// Add Card Data
/////////////////////////////////////////////////////////////////////////////////

// Standard I/O library.
#include <cstdio>

// Function prototype.
int GetCards(char vCardsInputed[]);
bool IsCardValue(char vC);

int AddCards()
{

	/////////////////////////////////////////////////////////////////////////////
	// 1. Get card data from user.
	char CardInputs[256];
	GetCards(CardInputs);
	printf("\nCards Entered: %s \n\n", CardInputs);

	/////////////////////////////////////////////////////////////////////////////
	// 2. Save card data to file.

	// Return success.
	return 0;
} // end AddCards
int GetCards(char vCardsInputed[])
{
	// Current input.
	char c = '\0';
	// Counter.
	short i = 0;
	// Need colour and value to form a card.
	bool IsNewCard = false;
	char CardValue = 0;
	char CardColour = 0;

	// Instructions.
	printf("\nEnter Random Card Data.\n");
	printf("B for black, R for red.\n");
	printf("2 to 9, T (10), J(ack), Q(ueen), K(ing), A(ce).\n");
	printf("\nEnter Cards: ");

	// Input data and validate.
	while ((c = getchar()) !=EOF)
	{
		// If lowercase, change to uppercase.
		if ((c >= 97) && (c <= 122))
			// Upper case.
			c -= 32;

		if (IsNewCard)
		{
			// Only accept card values.
			if (IsCardValue(c))
			{
				// Store value.
				CardValue = c;
				// Reset.
				IsNewCard = false;
			}
		}

		if ((c == 'B') || (c == 'R'))
		{
			if (IsNewCard) 
				// Ignore previous start.
				// Store value.
				CardColour = c;
			else
			{
				// Previous char not B or R.
				// Store value.
				CardColour = c;
				// Start of card.
				IsNewCard = true;
			}
		}

		// If there's a colour and value, add to string of cards.
		if ((CardColour) && (CardValue))
		{
			// Add character to string of card data.
			vCardsInputed[i++] = CardColour;
			vCardsInputed[i++] = CardValue;
			vCardsInputed[i++] = ' ';
			// Reset.
			CardColour = 0;
			CardValue = 0;
		}
	}

	// End with null char.
	vCardsInputed[i++] = '\0';

	return 0;
} // end GetCards
bool IsCardValue(char vC)
{
	// Return true if char is 2 to 9, T, J, Q, K or A.

	if ((vC >= '2') && (vC <= '9'))
		return true;
	else if ((vC == 'T'))
		return true;
	else if ((vC == 'J'))
		return true;
	else if ((vC == 'Q'))
		return true;
	else if ((vC == 'K'))
		return true;
	else if ((vC == 'A'))
		return true;
	else 
		return false;
} // end IsCardValue
