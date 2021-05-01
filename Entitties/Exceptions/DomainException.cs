using System;

namespace Gerenciadorxml.Entitties.Exceptions{

    class DomainException : ApplicationException{
        public DomainException(string message) : base(message) { 
        }
    }
}
