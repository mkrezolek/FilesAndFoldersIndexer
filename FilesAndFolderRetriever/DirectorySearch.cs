using System;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;

public class DirectorySearch
{
    public string rootPath;
    public string extention;
    

    public DirectorySearch(string selectedFolder)
	{
        
        this.rootPath = selectedFolder;
        TraverseDirectory();
	}

    public IEnumerable<FileInfo> TraverseDirectory()
    {
        Stack<DirectoryInfo> directoryStack = new Stack<DirectoryInfo>();
        directoryStack.Push(new DirectoryInfo(rootPath));
        while (directoryStack.Count > 0)
        {
            var dir = directoryStack.Pop();
            try
            {
                Parallel.ForEach (dir.GetDirectories(), item =>
                {
                    directoryStack.Push(item);
                });
            }
            catch (UnauthorizedAccessException)
            {
                continue; //skips the directory without an access
            }
            foreach (var f in dir.GetFiles())
            {
                yield return f;
            }
        }
    }
    
}

