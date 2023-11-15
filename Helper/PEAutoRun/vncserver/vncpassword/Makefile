all: clean vncpassword.exe

vncpassword.exe: main.c d3des.h d3des.c
	gcc -Os -o $@ -m64 main.c
	strip $@

clean:
	rm -f vncpassword.exe
