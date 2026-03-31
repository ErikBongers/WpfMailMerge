using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace WpfMailMerge;

internal class Cursor<T>(IEnumerator<T> enumerator)
    {
    private readonly IEnumerator<T> enumerator = enumerator;

    public void Eat() { this.MoveNext(); }
    public bool MoveNext() { return this.enumerator.MoveNext(); }
    public T Current { get { return this.enumerator.Current; } }
    public bool EOF() { return this.enumerator.Current is null; }

    public void Skip(Func<T, bool> until)
        {
        do
            {
            if (until(this.Current))
                return;
            }
        while (this.MoveNext());
        }
    }
