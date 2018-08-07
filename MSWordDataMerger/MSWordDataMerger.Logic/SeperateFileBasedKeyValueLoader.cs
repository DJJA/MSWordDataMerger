using System;
using System.Collections.Generic;
using System.IO;

namespace MSWordDataMerger.Logic
{
    class SeperateFileBasedKeyValueLoader : IKeyValueLoader
    {
        public KeyValuePairContainer LoadKeyValuePairs(string path)
        {
            var keyValuePairs = GetKeyValuePairs(path);
            var keyValuePairIterationHolders = new List<IterationKeyValuePairHolder>();

            var directories = Directory.GetDirectories(path);
            foreach (var directory in directories)                              // Get the interation directories, different iteration blocks 
            {
                var keyValuePairIterations = new List<List<KeyValuePair>>();
                var iterationDirectories = Directory.GetDirectories(directory);
                foreach (var iterationDirectory in iterationDirectories)        // Get the iteration directories, the same iteration block
                {
                    var pairs = GetKeyValuePairs(iterationDirectory);
                    if (pairs.Count > 0)                                         // If there are key value pairs add the iteration
                        keyValuePairIterations.Add(pairs);
                }
                if (keyValuePairIterations.Count > 0)                           // If there are key value pairs iterations, add it to the list
                {
                    var dirName = new DirectoryInfo(directory).Name;
                    var iterationKeyValuePairHolder = new IterationKeyValuePairHolder()
                    {
                        IterationBlockName = dirName,
                        KeyValuePairIterations = keyValuePairIterations
                    };
                    keyValuePairIterationHolders.Add(iterationKeyValuePairHolder);
                }
            }

            return new KeyValuePairContainer()
            {
                KeyValuePairs = keyValuePairs,
                KeyValuePairIterations = keyValuePairIterationHolders
            };
        }

        private List<KeyValuePair> GetKeyValuePairs(String path)
        {
            var keyValuePairs = new List<KeyValuePair>();
            var files = Directory.GetFiles(path);
            foreach (var file in files)
            {
                if (!Path.HasExtension(file))
                {
                    try
                    {
                        var fileContent = File.ReadAllText(file);
                        keyValuePairs.Add(new KeyValuePair(key: Path.GetFileName(file).Trim(), value: fileContent.Trim()));
                    }
                    catch (Exception e)
                    {
                        throw new MergerException($"Could not read value from key {Path.GetFileName(file).Trim()} located at \"{file}\"", e.Message);
                    }
                }
            }
            return keyValuePairs;
        }
    }
}
