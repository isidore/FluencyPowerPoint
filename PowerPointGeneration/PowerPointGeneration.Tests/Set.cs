using System;

namespace PowerPointGeneration.Tests
{
	public class Set
	{
		public SetCard[] cards;
		public Set(SetCard card1, SetCard card2, SetCard card3)
		{
			cards = new[] {card1, card2, card3};
		}

		
		public bool IsValidSet()
		{
			if (cards[0] == null || cards[1] == null || cards[2] == null)
			{
				return false;
			}
			return (IsFeatureValid(cards[0].number, cards[1].number, cards[2].number) &&
			        (IsFeatureValid(cards[0].symbol, cards[1].symbol, cards[2].symbol) &&
			         (IsFeatureValid(cards[0].shading, cards[1].shading, cards[2].shading) &&
			          IsFeatureValid(cards[0].color, cards[1].color, cards[2].color))));
		}

		/************************************************************************/

		public bool IsFeatureValid(int f1, int f2, int f3)
		{
			return (((f1 == f2) && (f1 == f3)) || ((f1 + f2 + f3) == 6));
		}
		public enum SetType
		{
			Invalid = -1, OneDifference=1, TwoDifferences=2, ThreeDifferences=3, FourDifferences=4
		}
		public SetType GetTypeOfSet()
		{
			if (!IsValidSet())
			{
				return SetType.Invalid;
			}
			int count = 0;
			Func<Func<SetCard, int>, int> differentCount = (w) => (w(cards[0]) == w(cards[1]) && w(cards[0]) == w(cards[2])) ? 0 : 1;
			count += differentCount(c => c.symbol);
			count += differentCount(c => c.color);
			count += differentCount(c => c.number);
			count += differentCount(c => c.shading);
			return (SetType) count;

		}
	}
}