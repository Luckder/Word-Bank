# Word-Bank
A simple python script to create a spreadsheet of words with their definitions and relevant information.

Uses the Oxford Dictionary API so you do need to have an API ID and Key from the Oxford Dictionary website. You can create an account for the free plan, but there is a limit of 1000 words.

The attached test.txt includes some words as examples to be used by the script, please keep your wordlist exactly like how you see it in the test.txt EXCEPT the last line where there are MULTIPLE words.

NOTES:
Do not have any lines with MULTIPLE words! The API will be unable to search for it in the dictionary and return an error. Hyphens/Dashes are still okay.
Please use only english words, no foreign languages, as the API only handles english in this script.
Unsure whether the differences between American English and British English will have an impact, but since the API is from Oxford and the mode in the script is fixed at "en-gb", please try to keep the words in British english.
Borrowed english words are still okay.
And lastly, please don't make any spelling mistakes! The API will be unable to find the word and you will get an error or even worse the wrong definition!

