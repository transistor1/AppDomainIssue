**The Problem**

When instantiating two, independent .NET COM-visible classes within the same, single-threaded COM client, .NET loads them both into the same AppDomain.

I am guessing that this is because they are being loaded into the same thread.

An example of this behavior is shown in [this GitHub repository](https://github.com/transistor1/AppDomainIssue).

Essentially, the demonstration is as follows:

1. Instantiate one COM class
1. Set an attribute on the first COM object which, in the back-end calls `SetData` on the `CurrentDomain`.
1. Instantiate a second, independent COM class (different interface name, GUIDs, etc)
1. Read the `AppDomain` attribute
1. Demonstrate that it appears the same
1. Also, get the hash code from both `AppDomain`s, noting that it is also the same

**Why is this a problem?**

When both classes have the `AppDomain.CurrentDomain.AssemblyResolve` event implemented (or any other AppDomain event, for that matter), the events can interfere with one another.  This is at least one complication; I am guessing that there may be others as well.

**An Idea**

I thought the best way of handling this would be to create a new AppDomain for each COM object.  Because I could not find (or Google) a way of doing this in a managed way, I thought it might be necessary to do it in unmanaged code.

I did a little detective work.  In OleView, the InprocServer32 attribute for a .NET COM-visible class is `mscoree.dll`.  So, I created a "shim" DLL which forwarded all of its `EXPORTS` to mscoree.dll.  By process of elimination (eliminating exports until the COM would no longer load), I discovered that [`DllGetClassObject`](https://msdn.microsoft.com/en-us/library/windows/desktop/ms680760(v=vs.85).aspx) in `mscoree` was responsible for starting up the .NET runtime, and returning the instantiated COM object.

So, what I can do is implement my own `DllGetClassObject`, like so:

1. Host the .NET runtime in an unmanaged assembly using [CLRCreateInstance](https://msdn.microsoft.com/en-us/library/dd233134(v=vs.110).aspx)
1. Create the object in a new `AppDomain`, and return it

(I'm guessing it's not as simple as it sounds, though)

**The Question**

Before I embark on this potentially difficult and lengthy process, I'd like to know:

1. Is there a managed way of getting a .NET COM-visible class to run in its own AppDomain?
1. If not, is this the "right" way of doing it, or am I missing an obvious solution?

This repository was created as a sample for [this StackOverflow question](http://stackoverflow.com/questions/41616334/how-to-instantiate-a-comvisible-class-into-its-own-appdomain-in-a-single-threade).
