using System;
 
    public class LetterAdd
    {
        private char letter;
        private int num; 
        public LetterAdd(char c)
        {
            letter = c;
            num = 0;
        }

        public bool addToNum()
        {
            if (num >= 9)
            {
                num = 0;
                return false;
            }
            
                num += 1;
            return true;           

        }

    }

