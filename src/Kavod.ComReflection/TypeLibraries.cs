using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;

namespace Kavod.ComReflection
{
    public class TypeLibraries
    {
        private readonly IList<TypeLibrary> _loadedLibraries = new List<TypeLibrary>();

        public IEnumerable<TypeLibrary> LoadedLibraries => _loadedLibraries;

        public TypeLibrary LoadLibrary(LibraryRegistration registration)
        {
            var library = LoadedLibraries.FirstOrDefault(l => l.FilePath == registration.FilePath);
            if (library == null)
            {
                library = new TypeLibrary(registration.FilePath, this);
                _loadedLibraries.Add(library);
            }
            return library;
        }

        internal TypeLibrary LoadLibrary(ITypeLib typeLib)
        {
            var name = ComHelper.GetTypeLibName(typeLib);
            var library = LoadedLibraries.FirstOrDefault(l => l.Name == name);
            if (library == null)
            {
                library = new TypeLibrary(typeLib, this);
                _loadedLibraries.Add(library);
            }
            return library;
        }

        internal TypeLibrary LoadLibrary(Guid libGuid)
        {
            var library = LoadedLibraries.FirstOrDefault(l => l.Guid.Equals(libGuid));
            if (library == null)
            {
                var reg = LibraryRegistration.GetComTypeRegistryEntry(libGuid);
                library = new TypeLibrary(reg.FilePath, this);
                _loadedLibraries.Add(library);
            }
            return library;
        }
    }
}
