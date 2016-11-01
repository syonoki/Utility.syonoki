using System;
using System.Collections.Generic;

namespace Utility.syonoki.ExtensionMethods {
    public static class DictionaryExtension {

        public static void ForEach<TKey, TValue>(this Dictionary<TKey, TValue> dict, Action<TKey, TValue> action) {
            if (dict == null) 
                throw new ArgumentNullException("argument dict is null");
            
            if (action == null) 
                throw new ArgumentNullException("argument action is null");

            foreach (KeyValuePair<TKey, TValue> pair in dict) {
                action(pair.Key, pair.Value);
            }
        }
    }
}