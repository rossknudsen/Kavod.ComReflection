using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;

namespace Kavod.ComReflection
{
    public class TypeLibraries
    {
        private const string StdOleLibGuidString = "{00020430-0000-0000-C000-000000000046}";
        internal static readonly Guid StdOleLibGuid = new Guid(StdOleLibGuidString);

        private readonly ObservableCollection<TypeLibrary> _libraries;

        public TypeLibraries()
        {
            _libraries = new ObservableCollection<TypeLibrary>();
            LoadedLibraries = new ReadOnlyObservableCollection<TypeLibrary>(_libraries);
        }

        public ReadOnlyObservableCollection<TypeLibrary> LoadedLibraries { get; }

        public TypeLibrary LoadLibrary(LibraryRegistration registration)
        {
            var library = LoadedLibraries.FirstOrDefault(l => l.FilePath == registration.FilePath);
            if (library == null)
            {
                library = new TypeLibrary(registration.FilePath, this);
                _libraries.Add(library);
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
                _libraries.Add(library);
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
                _libraries.Add(library);
            }
            return library;
        }
    }
}
