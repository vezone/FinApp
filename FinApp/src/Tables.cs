using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace FinApp.src
{
    class Tables
    {
        public int TablesCount => Content.Count;
        public List<Table> Content { get; }

        public Tables(int count)
        {
            Content = new List<Table>(count);
        }

        public Tables(Tables value)
        {
            Content = new List<Table>(value.Content);
        }

        public Table this[int index]
        {
            get
            {
                return Content[index];
            }
            set
            {
                if (index >= 0 && index < Content.Count)
                {
                    Content[index] = new Table(value);
                }
                else
                {
                    Content.Add(value);
                }
            }
        }

        public void Add(Table value)
        {
            Content.Add(value);
        }

        public bool Remove(int index)
        {
            if (index < Content.Count
                &&
                index >= 0)
            {
                Content.RemoveAt(index);
                return true;
            }
            return false;
        }

        public int NumberOfIndividDates()
        {
            int length = -1;
            List<Table> tList = new List<Table>(Content);

            for (int i = 0; i < tList.Count; i++, length++)
            {
                for (int j = i + 1; j < tList.Count; j++)
                {
                    if (tList[i].FirstName == tList[j].FirstName)
                    {
                        tList.RemoveAt(j);
                    }
                }
                tList.RemoveAt(i);
                i--;
            }
            return length;
        }

        public List<ParsedTables> ParseTablesFirstName()
        {
            List<ParsedTables> plist = new List<ParsedTables>(1);
            List<Table> tList = new List<Table>(Content);

            for (int i = 0; i < tList.Count; i++)
            {
                ParsedTables pt = new ParsedTables(tList[i].FirstName);
                pt.Add(tList[i]);
                for (int j = i + 1; j < tList.Count; j++)
                {
                    if (tList[i].FirstName == tList[j].FirstName)
                    {
                        pt.Add(tList[j]);
                        tList.RemoveAt(j);
                        j--;
                    }
                }
                plist.Add(pt);
            }
            return plist;
        }
        public List<ParsedTables> ParseTablesSecondName()
        {
            List<ParsedTables> plist = new List<ParsedTables>(1);
            List<Table> tList = new List<Table>(Content);

            for (int i = 0; i < tList.Count; i++)
            {
                ParsedTables pt = new ParsedTables(tList[i].SecondName);
                pt.Add(tList[i]);
                for (int j = i + 1; j < tList.Count; j++)
                {
                    if (tList[i].SecondName == tList[j].SecondName)
                    {
                        pt.Add(tList[j]);
                        tList.RemoveAt(j);
                        j--;
                    }
                }
                plist.Add(pt);
            }
            return plist;
        }

        public List<ParsedTables> ParseTablesByDates()
        {
            List<Table> tList =
                new List<Table>(Content);
            List<ParsedTables> m_ParsedTables =
                new List<ParsedTables>(NumberOfIndividDates());

            for (int i = 0; i < tList.Count; i++)
            {
                ParsedTables parsedTables = new ParsedTables(tList[i].FirstName);
                parsedTables.Add(tList[i]);
                for (int j = i + 1; j < tList.Count; j++)
                {
                    if (tList[i].FirstName == tList[j].FirstName)
                    {
                        parsedTables.Add(tList[j]);
                        tList.RemoveAt(j);
                        j--;
                    }
                }
                m_ParsedTables.Add(parsedTables);
            }
            return m_ParsedTables;
        }

        public delegate List<ParsedTables> ParsingFunction();

        public (List<string> names, List<double> values)
                SplitList(ParsingFunction parsingFunction)
        {
            List<string> names    = new List<string>();
            List<double> values   = new List<double>();
            List<ParsedTables> pt = parsingFunction();

            for (int i = 0; i < pt.Count; i++)
            {
                names .Add(pt[i].Name);
                values.Add(pt[i].GetCacheSpending());
            }

            return (names, values);
        }

    }
}
