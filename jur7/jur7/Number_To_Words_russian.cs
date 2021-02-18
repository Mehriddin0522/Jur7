using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace jur7
{
    class Number_To_Words_russian
    {
        public int DG_POWER = 10;

        private String[,] a_power = new String[,]{
     {"0", ""            , ""             ,""              },  // 1
     {"1", "тысяча "     , "тысячи "      ,"тысяч "        },  // 2
     {"0", "миллион "    , "миллиона "    ,"миллионов "    },  // 3
     {"0", "миллиард "   , "миллиарда "   ,"миллиардов "   },  // 4
     {"0", "триллион "   , "триллиона "   ,"триллионов "   },  // 5
     {"0", "квадриллион ", "квадриллиона ","квадриллионов "},  // 6
     {"0", "квинтиллион ", "квинтиллиона ","квинтиллионов "},  // 7
     {"0", "секстиллион ", "секстиллиона ","секстиллионов "},  // 8
     {"0", "септиллион " , "септиллиона " ,"септиллионов " },  // 9
     {"0", "октиллион "  , "октиллиона "  ,"октиллионов "  },  // 10
    };

        private String[,] digit = new String[,] {
     {""       ,""       , "десять "      , ""            ,""          },
     {"один "  ,"одна "  , "одиннадцать " , "десять "     ,"сто "      },
     {"два "   ,"две "   , "двенадцать "  , "двадцать "   ,"двести "   },
     {"три "   ,"три "   , "тринадцать "  , "тридцать "   ,"триста "   },
     {"четыре ","четыре ", "четырнадцать ", "сорок "      ,"четыреста "},
     {"пять "  ,"пять "  , "пятнадцать "  , "пятьдесят "  ,"пятьсот "  },
     {"шесть " ,"шесть " , "шестнадцать " , "шестьдесят " ,"шестьсот " },
     {"семь "  ,"семь "  , "семнадцать "  , "семьдесят "  ,"семьсот "  },
     {"восемь ","восемь ", "восемнадцать ", "восемьдесят ","восемьсот "},
     {"девять ","девять ", "девятнадцать ", "девяносто "  ,"девятьсот "}
    };

        public String toWords(double sum)
        {
            int TAUSEND = 1000;
            int i, mny;
            System.Text.StringBuilder result = new System.Text.StringBuilder("");
            double divisor; // делитель
            double psum = sum;

            int one = 1;
            int four = 2;
            int many = 3;

            int hun = 4;
            int dec = 3;
            int dec2 = 2;

            if (sum == 0)
                return "ноль ";
            if (sum.CompareTo(0) < 0)
            {
                result.Append("минус ");
                psum = psum * (-1);
            }

            for (i = 0, divisor = 1; i < DG_POWER - 1; i++)
            {
                divisor = divisor * (TAUSEND);
                if (sum.CompareTo(divisor) < 0)
                {
                    i++;
                    break; // no need to go further
                }
            }
            // start from previous value
            for (; i >= 0; i--)
            {
                mny = (int)(psum / (divisor));
                psum = psum % (divisor);
                // str="";
                if (mny == 0)
                {
                    // if(i>0) continue;
                    if (i == 0)
                    {
                        result.Append(a_power[i, one]);
                    }
                }
                else
                {
                    if (mny >= 100)
                    {
                        result.Append(digit[mny / 100, hun]);
                        mny %= 100;
                    }
                    if (mny >= 20)
                    {
                        result.Append(digit[mny / 10, dec]);
                        mny %= 10;
                    }
                    if (mny >= 10)
                    {
                        result.Append(digit[mny - 10, dec2]);
                    }
                    else
                    {
                        if (mny >= 1)
                            result.Append(digit[mny, "0".Equals(a_power[i, 0]) ? 0
                                : 1]);
                    }
                    switch (mny)
                    {
                        case 1:
                            result.Append(a_power[i, one]);
                            break;
                        case 2:
                        case 3:
                        case 4:
                            result.Append(a_power[i, four]);
                            break;
                        default:
                            result.Append(a_power[i, many]);
                            break;
                    }
                }
                divisor = divisor / (TAUSEND);
            }
            return result.ToString();
        }
    }
}
