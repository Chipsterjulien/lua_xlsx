# Binaire Babet local

Place ici le binaire Babet utilisé pour les tests :

```text
bin/babet
```

Puis rends-le exécutable si nécessaire :

```sh
chmod +x bin/babet
./run_tests.sh
```

`run_tests.sh` cherche les runtimes dans cet ordre :

1. `BABET_BIN`, lorsqu'elle est définie ;
2. `bin/babet` ;
3. `babet` dans le `PATH`.

Le binaire lui-même n'est pas versionné par ce projet.
