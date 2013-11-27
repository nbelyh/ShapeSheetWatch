template<class T>
class Ptr
{
public:
	Ptr() :p(), c() {}
	explicit Ptr(T* s) :p(s), c(new unsigned(1)) {}

	Ptr(const Ptr& s) :p(s.p), c(s.c) { if(c) ++*c; }

	Ptr& operator=(const Ptr& s) 
	{ if(this!=&s) { clear(); p=s.p; c=s.c; if(c) ++*c; } return *this; }

	~Ptr() { clear(); }

	void clear() 
	{ 
		if(c)
		{
			if(*c==1) delete p; 
			if(!--*c) delete c; 
		} 
		c=0; p=0; 
	}

	T* get() const { return (c)? p: 0; }
	T* operator->() const { return get(); }
	T& operator*() const { return *get(); }

private:
	T* p;
	unsigned* c;
};