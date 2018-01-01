# extract_content
One library for Visual Basic that you can extract any string from a string (e.g. a txt/html file) that is located between 2 known strings or characters.

In this dll you can find 4 functions.
1. ExtractContentWithoutKeyWords: with this function you get the string with the search words that are between them.
  e.g.         Dim strInput, strLookFor, strLookFor2, strToExtract As String
               Dim ec As New ExtractContent.ExtractContent
               strInput = "Hello word! Have a Nice Year!!!"
               strLookFor = "! "
               strLookFor2 = "!!!"
               strToExtract = ec.ExtractContentWithoutKeyWords(strInput, strLookFor, strLookFor2)
         In this example, strToExtract will be "Have a Nice Year"
2.  ExtractContentWith1stKeyWord: with this function you get the string that you looking for with the first searching string.
  e.g.         Dim strInput, strLookFor, strLookFor2, strToExtract As String
               Dim ec As New ExtractContent.ExtractContent
               strInput = "Hello word! Have a Nice Year!!!"
               strLookFor = "Have"
               strLookFor2 = "!!"
               strToExtract = ec.ExtractContentWith1stKeyWord(strInput, strLookFor, strLookFor2)
         In this example, strToExtract will be "Have a Nice Year!"
3. ExtractContentWith2ndKeyWord: with this function you get the string that you looking for with the second searching string.
  e.g.         Dim strInput, strLookFor, strLookFor2, strToExtract As String
               Dim ec As New ExtractContent.ExtractContent
               strInput = "Hello word! Have a Nice Year!!!"
               strLookFor = "! "
               strLookFor2 = "Year!"
               strToExtract = ec.ExtractContentWith2ndKeyWord(strInput, strLookFor, strLookFor2)
         In this example, strToExtract will be "Have a Nice Year!"
4. ExtractContentWithKeyWords: with this function, you get the string that you looking for, with both the searching strings that you refer.
  e.g.         Dim strInput, strLookFor, strLookFor2, strToExtract As String
               Dim ec As New ExtractContent.ExtractContent
               strInput = "Hello word! Have a Nice Year!!!"
               strLookFor = "Have "
               strLookFor2 = "Year!"
               strToExtract = ec.ExtractContentWithKeyWords(strInput, strLookFor, strLookFor2)
         In this example, strToExtract will be "Have a Nice Year!"
