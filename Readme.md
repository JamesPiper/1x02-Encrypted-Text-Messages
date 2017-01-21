# Encrypted Text Messages

OVERVIEW

Encrypt a text message using a one-time key.

CONCEPT
· A one-time pad or key provides a secure method of encrypting a message provided the key is secured, random and not reused.
· The key is the same length of the message.
· One-time pads rely on synchronous keys. Sender and receiver use the same key to encrypt and decrypt the message. This means this message is not appropriate for broadcasting a message to more than one person.
· The method is designed to secure the message while en route and not as a means for storing a message.
· Keys that are reused or compromised leaves the message vulnerable to cracking.
· Strict procedures are required for the safe use of this method. Human error in following these procedures results in vulnerable messages.

OBJECTIVE
· Implement a one-time pad process on a short text message.
· A user enters both a plaintext message and key to create a cyphertext message.
· A receiver decrypts the cyphertext using the key.
· The message would be limited to 180 characters and thus would be appropriate for SMS, Twitter etc.
· Only uppercase letters of the Latin alphbet are used in the cyphertext (i.e, A to Z). Spaces, punctuation marks etc in the plaintext are ignored.

LIMITATIONS
· The creation of a random key. Non-random keys could result in successful brute force attacks or letter frequency attacks.
· Security of the key. Anyone with the key could decrypt the message.
· Loss of the key or an inaccurate key results in lost or unreadable messages.
· Appropriate for short messages.
· Synchronizing keys between sender and receiver.
· Text message limited to 26 letters of Latin alphabet.
· Time consuming to generate key, sync keys and to encrypt and decrypt.






 

