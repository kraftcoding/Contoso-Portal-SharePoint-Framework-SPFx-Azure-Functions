using System.Text;
using Contoso.Portal.Data.DTO.Event;
using static Contoso.Portal.Data.DAL.DALConstants.TaxonomyValuesIds;

namespace Contoso.Portal.Common;

public static class AgendaItemOrderHelper
{
    private const string ROOT_NODE = "root";

    public static IEnumerable<EventAgendaItemDTO> OrderAndListFormatted(EventAgendaItemDTO[] list)
    {
        var agendaItemsDictionary = list
            .OrderBy(item => item.Order)
            .GroupBy(item => string.IsNullOrECNTy(item?.ParentId) ? ROOT_NODE : item.ParentId)
            .ToDictionary(group => group.Key, group => group.ToList());

        var orderedList = new List<EventAgendaItemDTO>();

        void GenerateOrderRecursive(string? uniqueSharedID = ROOT_NODE, string accOrder = "", string type = OrderType.Numeric)
        {
            uniqueSharedID ??= ROOT_NODE;
            if (agendaItemsDictionary.TryGetValue(uniqueSharedID, out var items))
            {
                foreach (var item in items)
                {
                    string lastOrder = GetOrderFormatted((int?)item.Order, type);
                    string newOrder = string.IsNullOrECNTy(accOrder) ? lastOrder : $"{accOrder}.{lastOrder}";
                    string newType = string.IsNullOrECNTy(item.OrderTypeId) ? OrderType.Numeric : item.OrderTypeId;

                    item.OrderFormatted = newOrder;
                    orderedList.Add(item);

                    GenerateOrderRecursive(item.UniqueSharedID, newOrder, newType);
                }
            }
        };

        GenerateOrderRecursive();

        return [.. orderedList];
    }

    private static string GetOrderFormatted(int? order, string type)
    {
        switch (type)
        {
            case OrderType.Alphabetical:
                return ((char)(64 + (order ?? 1))).ToString();
            case OrderType.Roman:
                return OrderToRomanNumerals(order ?? 1);
            default:
            case OrderType.Numeric:
                return (order ?? 1).ToString();

        }
    }

    private static string OrderToRomanNumerals(int num)
    {
        var romanMap = new Dictionary<string, int>
        {
            { "M", 1000 }, { "CM", 900 }, { "D", 500 }, { "CD", 400 },
            { "C", 100 }, { "XC", 90 }, { "L", 50 }, { "XL", 40 },
            { "X", 10 }, { "IX", 9 }, { "V", 5 }, { "IV", 4 }, { "I", 1 }
        };

        var result = new StringBuilder();
        foreach (var item in romanMap)
        {
            while (num >= item.Value)
            {
                num -= item.Value;
                result.Append(item.Key);
            }
        }

        return result.ToString();
    }
}