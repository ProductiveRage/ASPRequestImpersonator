using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using ASPRequestImpersonator;

namespace Examples
{
    [ComVisible(true)]
    [ProgId("ASPRequestImpersonator.ExampleGenerator")]
    [Guid("05E6FBDE-2AFD-4466-A2C4-780D0E046BF7")]
    public class ExampleGenerator
    {
        public LateBindingComWrapper GetExample1()
        {
            var formData = new[]
            {
                new[] { "Key1", "Value1 with spaces from Form" },
                new[] { "Key4", "Value5" }
            };
            
            var querystringData = new[]
            {
                new[] { "Key1", "Value1 with spaces from Querystring" },
                new[] { "Key2", "Value2", "Value3" },
                new[] { "Key3", "Value4" }
            };
            
            var serverVariablesData = new string[0][];

            return new LateBindingComWrapper(
                new RequestImpersonator(
                    GetData(formData),
                    GetData(querystringData),
                    GetData(serverVariablesData)
                )
            );
        }

        private RequestDataSnapshot GetData(params string[][] values)
        {
            if (values == null)
                throw new ArgumentNullException("value");
            if (values.Any(v => v == null))
                throw new ArgumentException("Null entry enountered in values");

            var lines = new List<KeyValuePair<string, RequestStringList>>();
            foreach (var value in values)
                lines.Add(GetLineEntry(value));
            return new RequestDataSnapshot(lines);
        }

        private KeyValuePair<string, RequestStringList> GetLineEntry(params string[] values)
        {
            if (values == null)
                throw new ArgumentNullException("values");
            if (values.Any(v => v == null))
                throw new ArgumentException("Null entry enountered in values");
            if (!values.Any())
                throw new ArgumentException("There must be at least one value specified, for the entry key");

            var key = values.First().Trim();
            if (key == null)
                throw new ArgumentException("First value may not be empty");

            return new KeyValuePair<string, RequestStringList>(
                key,
                new RequestStringList(values.Skip(1))
            );
        }
    }
}
