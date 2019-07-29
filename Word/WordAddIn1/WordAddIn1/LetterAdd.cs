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
    public int getNum()
    {
        return num;
    }
    public void addToNum()
    {
        num += 1;
        if (num >= 10)
        {
            num = 0; 
        }

         

    }

}

