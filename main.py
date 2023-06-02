

print("====================");
print("Matura data analyzer");
print("====================");

i = True
while i:
    to_read_file = input("Podaj ścieżke do pliku: ");
    if to_read_file.split(".")[1] == "xls" or to_read_file.split(".")[1] == "xlsx":
        print("Prawidłowy format")
        i = False
    else: 
        print("Nieprawidłowy format pliku!")