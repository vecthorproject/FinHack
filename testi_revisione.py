# -*- coding: utf-8 -*-
"""
Testi del report e della presentazione nello stile della revisione di settembre.

La revisione ha riscritto quasi tutti i commenti agli indicatori seguendo alcuni
criteri costanti:
- il percorso 2021-2024 si racconta con gli anni e i valori dei punti di svolta;
- la dinamica dell'impresa si confronta sempre con quella della mediana di
  settore, anno per anno quando le due divergono;
- per i margini si dice di quanti punti percentuali e' cambiato il valore
  nell'ultimo esercizio;
- per gli indici con soglia (struttura, liquidita') si dice se e da quando il
  valore sta sopra l'unita';
- niente giudizi che il dato non sostiene.

Qui quei criteri diventano regole sui dati, cosi' valgono per qualunque azienda:
sull'export di prova i testi coincidono con quelli scritti a mano nella
revisione, su un'azienda con un andamento diverso escono le frasi adatte a
quell'andamento.
"""

import math

from report_corp import con_articolo, format_euro

ANNI = ('2021', '2022', '2023', '2024')

# Soglie di lettura.
LIEVE = 5.0          # % : sotto, una variazione e' "lieve"
MARCATO = 50.0       # % : sopra, un salto annuo e' "molto marcato"
FORTE = 30.0         # % : sopra, un miglioramento sul periodo e' "forte"
VICINO = 10.0        # % : entro, due valori sono "prossimi"


# ----------------------------------------------------------------- numeri ----

def _ok(x):
    return x is not None and not (isinstance(x, float) and math.isnan(x))


def n(x, u=''):
    """Numero all'italiana con l'unita': "5,43%", "1,88"."""
    return f"{format_euro(x)}{u}" if _ok(x) else 'n.d.'


def na(x, u='', prep=None):
    """Numero con l'articolo concordato: "al 5,43%", "dall'8,94%", "lo 0,96"."""
    return f"{con_articolo(format_euro(x), prep)}{u}" if _ok(x) else 'n.d.'


def pp(x):
    """Differenza fra due percentuali, in punti percentuali."""
    return f"{format_euro(abs(x))} punti percentuali"


def var_pct(da, a):
    if not (_ok(da) and _ok(a)) or abs(da) < 1e-9:
        return None
    return (a - da) / abs(da) * 100


# ----------------------------------------------------------------- forma ----

class Serie:
    """Una serie 2021-2024 letta come la legge chi commenta."""

    def __init__(self, valori):
        # Arrotondati come vengono stampati: 0,001 e 0,004 si leggono entrambi 0,00, e il
        # testo non puo' dire che l'indicatore e' salito.
        self.punti = [(int(a), round(float(valori[a]), 2)) for a in ANNI if _ok(valori.get(a))]
        self.valida = len(self.punti) >= 2
        # anno e valore piu' recenti servono anche quando manca lo storico
        self.aN, self.vN = self.punti[-1] if self.punti else (2024, None)
        if not self.valida:
            return
        (self.a0, self.v0), (self.aN, self.vN) = self.punti[0], self.punti[-1]
        self.aP, self.vP = self.punti[-2]                      # l'anno prima dell'ultimo
        self.amax, self.vmax = max(self.punti, key=lambda p: p[1])
        self.amin, self.vmin = min(self.punti, key=lambda p: p[1])
        interni = [a for a, _ in self.punti[1:-1]]
        passi = [b - a for (_, a), (_, b) in zip(self.punti, self.punti[1:])]
        self.sale_sempre = all(p > 0 for p in passi)
        self.scende_sempre = all(p < 0 for p in passi)
        self.picco = (self.amax in interni and self.vmax > self.v0 and self.vmax > self.vN)
        self.minimo = (self.amin in interni and self.vmin < self.v0 and self.vmin < self.vN)
        self.var = var_pct(self.v0, self.vN)
        self.var_ultimo = var_pct(self.vP, self.vN)

    @property
    def completa(self):
        """Copre tutti gli anni dal primo all'ultimo: solo allora e' un quadriennio."""
        return self.valida and self.a0 == int(ANNI[0]) and self.aN == int(ANNI[-1])

    def _sale_fino_a(self, anno):
        tratto = [v for a, v in self.punti if a <= anno]
        return all(b > a for a, b in zip(tratto, tratto[1:]))

    def valore(self, anno):
        for a, v in self.punti:
            if a == int(anno):
                return v
        return None

    def sempre_sopra(self, soglia):
        return all(v > soglia for _, v in self.punti)

    def primo_anno_sopra(self, soglia):
        """Primo anno da cui il valore resta sopra la soglia fino alla fine."""
        anno = None
        for a, v in self.punti:
            if v > soglia:
                anno = anno or a
            else:
                anno = None
        return anno


def sopra_in_tutti(az, sett):
    """Anni in cui l'impresa sta sopra la mediana, e quelli in cui no."""
    sopra, sotto = [], []
    for a, v in az.punti:
        m = sett.valore(a)
        if m is None:
            continue
        (sopra if v > m else sotto).append(a)
    return sopra, sotto


# ------------------------------------------------------ percorso impresa ----

def percorso_margine(az, variante):
    """Il percorso di un margine nelle tre forme usate nella revisione."""
    if not az.valida:
        return ''
    u = '%'
    if az.picco:
        if variante == 'ebitda':
            return (f"{nel_periodo(az)} l'indicatore aziendale passa {na(az.v0, u, 'da')} del {az.a0} "
                    f"{na(az.vmax, u, 'a')} del {az.amax}, per poi scendere nel {az.aN} {na(az.vN, u, 'a')}")
        if variante == 'ebit':
            return (f"Dopo essere salito {na(az.v0, u, 'da')} del {az.a0} {na(az.vmax, u, 'a')} del "
                    f"{az.amax}, l'indicatore scende {na(az.vN, u, 'a')} nel {az.aN}")
        return (f"L'indicatore aziendale migliora {na(az.v0, u, 'da')} del {az.a0} {na(az.vmax, u, 'a')} "
                f"del {az.amax}, ma si riduce {na(az.vN, u, 'a')} nel {az.aN}")
    if az.minimo:
        return (f"Dopo essere sceso {na(az.v0, u, 'da')} del {az.a0} {na(az.vmin, u, 'a')} del "
                f"{az.amin}, l'indicatore risale {na(az.vN, u, 'a')} nel {az.aN}")
    verbo = 'sale' if az.vN > az.v0 else ('scende' if az.vN < az.v0 else 'resta')
    if verbo == 'resta':
        return f"{nel_periodo(az)} l'indicatore aziendale resta sui livelli del {az.a0} ({n(az.vN, u)})"
    return (f"{nel_periodo(az)} l'indicatore aziendale {verbo} {na(az.v0, u, 'da')} del {az.a0} "
            f"{na(az.vN, u, 'a')} del {az.aN}")


def nel_periodo(s):
    return "Nel quadriennio" if s.completa else f"Tra il {s.a0} e il {s.aN}"


def tutto_il_periodo(s, prep):
    return f"{prep} tutto il quadriennio" if s.completa else "in tutti gli anni disponibili"


def periodo_anni(s):
    return f"Nel {'quadriennio' if s.completa else 'periodo'} {s.a0}-{s.aN}"


def passo_per_passo(az, u=''):
    """
    Serie a zig-zag raccontata anno per anno:
    "dopo essere sceso dal 200,14% del 2021 al 161,23% del 2022, e' risalito al
    179,01% nel 2023 e si e' nuovamente ridotto nel 2024".
    L'ultimo valore non si ripete: lo dice gia' la frase che precede.
    """
    passi = list(zip(az.punti, az.punti[1:]))
    if not passi:
        return ''
    pezzi, direzioni = [], []
    inizio_fermo = None                     # primo anno del tratto in cui il valore non cambia
    for i, ((aa, va), (ab, vb)) in enumerate(passi):
        ultimo = i == len(passi) - 1 and i > 0
        if vb == va:
            if inizio_fermo is not None:    # il tratto fermo continua: si allunga, non si ripete
                pezzi.pop()
            else:
                inizio_fermo = aa
            if not pezzi:
                pezzi.append(f"dopo essere rimasto {na(vb, u, 'a')} dal {inizio_fermo} al {ab}")
            elif ultimo:
                pezzi.append(f"\u00e8 rimasto invariato fino al {ab}" if inizio_fermo != aa else
                             f"\u00e8 rimasto invariato nel {ab}")
            else:
                pezzi.append(f"\u00e8 rimasto {na(vb, u, 'a')} fino al {ab}")
            continue
        inizio_fermo = None
        giu = vb < va
        if not pezzi:
            pezzi.append(f"dopo essere {'sceso' if giu else 'salito'} {na(va, u, 'da')} del {aa} "
                         f"{na(vb, u, 'a')} del {ab}")
            direzioni.append(giu)
            continue
        cambia = bool(direzioni) and giu != direzioni[-1]
        gia_visto = giu in direzioni
        if not direzioni:                   # primo movimento dopo un tratto fermo
            verbo = "si \u00e8 ridotto" if giu else "\u00e8 salito"
        elif giu:
            verbo = ("si \u00e8 nuovamente ridotto" if cambia and gia_visto else
                     "si \u00e8 ridotto" if cambia else "\u00e8 ulteriormente sceso")
        else:
            verbo = ("\u00e8 nuovamente salito" if cambia and gia_visto else
                     "\u00e8 risalito" if cambia else "\u00e8 ancora salito")
        if ultimo and pezzi[-1].startswith(verbo + " "):
            pezzi.append(f"di nuovo nel {ab}")      # "e' ancora salito ... e di nuovo nel 2024"
        else:
            pezzi.append(f"{verbo} nel {ab}" if ultimo else f"{verbo} {na(vb, u, 'a')} nel {ab}")
        direzioni.append(giu)
    # Il primo pezzo e' la subordinata ("dopo essere ..."): la principale segue dopo la virgola.
    testa, resto = pezzi[0], pezzi[1:]
    if not resto:                           # senza principale la frase resterebbe monca
        return ''
    if len(resto) == 1:
        return testa + ", " + resto[0]
    congiunzione = " ed " if resto[-1].startswith("\u00e8") else " e "
    return testa + ", " + ", ".join(resto[:-1]) + congiunzione + resto[-1]


def movimento_settore(sett, u=''):
    """Com'e' andata la mediana di settore nel periodo."""
    if not sett.valida:
        return ''
    if sett.sale_sempre:
        return f"continua a crescere in tutto il periodo passando {na(sett.v0, u, 'da')} {na(sett.vN, u, 'a')}"
    if sett.scende_sempre:
        return f"diminuisce in modo continuo, {na(sett.v0, u, 'da')} nel {sett.a0} {na(sett.vN, u, 'a')} nel {sett.aN}"
    verbo = 'sale' if sett.vN > sett.v0 else 'scende'
    return f"{verbo} {na(sett.v0, u, 'da')} {na(sett.vN, u, 'a')}"


# ------------------------------------------------ bullet del Benchmark ----
# Nel capitolo del posizionamento ogni indicatore ha un bullet: valore 2024 contro
# la mediana, percorso, confronto con la dinamica del settore.

def _testa(nome_art, az_val, med, u):
    if not _ok(az_val):
        return f"• {nome_art} non risulta disponibile per il 2024."
    base = f"• {nome_art} è pari {na(az_val, u, 'a')}"
    if _ok(med):
        base += f", contro {na(med, u)} della mediana settoriale"
    return base + "."


def bullet_margine(chiave, nome_art, az, sett):
    """EBITDA, EBIT e Margine di Profitto."""
    u = '%'
    frasi = [_testa(nome_art, az.valore(2024), sett.valore(2024), u)]
    if not az.valida:
        return frasi[0]
    percorso = percorso_margine(az, chiave)

    if chiave == 'ebitda':
        # "...; nello stesso periodo il settore sale dall'8,94% al 10,43%."
        verbo = 'sale' if sett.valida and sett.vN > sett.v0 else 'scende'
        frase = percorso
        if sett.valida:
            frase += f"; nello stesso periodo il settore {verbo} {na(sett.v0, u, 'da')} {na(sett.vN, u, 'a')}"
        frasi.append(frase + ".")
        if az.picco and az.amax == az.aP:
            gap_prima = (sett.valore(az.aP) - az.vP) if sett.valore(az.aP) is not None else None
            gap_ora = sett.vN - az.vN if sett.valida else None
            coda = ''
            if gap_prima is not None and gap_ora is not None and gap_ora > 0:
                coda = (" e amplia il divario rispetto al comparto" if gap_ora > gap_prima
                        else " pur riducendo il divario rispetto al comparto")
            frasi.append(f"Il {az.aN} interrompe quindi il miglioramento osservato fino al {az.amax}{coda}.")
        if sett.valida and az.vN < sett.vN:
            frasi.append(f"La distanza dal settore è di {pp(sett.vN - az.vN)}.")
        return " ".join(frasi)

    if chiave == 'ebit':
        frase = percorso + "."
        frasi.append(frase)
        if az.var_ultimo is not None and az.vN < az.vP:
            precedenti = [v for a, v in az.punti if az.aN - 2 <= a < az.aN]
            sotto_due = len(precedenti) == 2 and all(az.vN < v for v in precedenti)
            f = f"La contrazione di {pp(az.vP - az.vN)} rispetto al {az.aP}"
            if sotto_due:
                f += " riporta il margine al di sotto dei livelli osservati nei due anni precedenti"
            if sett.valida:
                f += f", mentre la mediana settoriale {movimento_settore(sett, u)}"
            frasi.append(f + ".")
        elif sett.valida:
            frasi.append(f"Nello stesso periodo la mediana settoriale {movimento_settore(sett, u)}.")
        return " ".join(frasi)

    # Margine di Profitto
    frasi.append(percorso + ".")
    if az.picco and az.amax == az.aP:
        frasi.append(f"La dinamica conferma che il recupero di redditività registrato fino al {az.amax} "
                     f"non si è consolidato nell'ultimo esercizio.")
    if sett.valida:
        stesso_picco = sett.picco and sett.amax == az.amax
        if az.picco and stesso_picco:
            frasi.append(f"L'impresa replica il trend di settore, in crescita fino al {sett.amax} (con un "
                         f"margine di profitto mediano salito {na(sett.v0, u, 'da')} {na(sett.vmax, u, 'a')}) "
                         f"e in diminuzione nell'ultimo anno, con un valore {na(sett.vN, u, 'di')}.")
        else:
            frasi.append(f"Nello stesso periodo la mediana di settore {movimento_settore(sett, u)}.")
    return " ".join(frasi)


def bullet_struttura1(nome_art, az, sett):
    frasi = [_testa(nome_art, az.valore(2024), sett.valore(2024), '')]
    if not az.valida:
        return frasi[0]
    if az.vN > 1:
        frasi.append("Il valore superiore all'unità indica che il patrimonio netto copre integralmente "
                     "le immobilizzazioni.")
    else:
        frasi.append("Il valore inferiore all'unità indica che il patrimonio netto non basta a coprire "
                     "le immobilizzazioni.")
    frasi.append(_evoluzione_struttura1(az, sett) + ".")
    if sett.valida and az.var and sett.vN > sett.v0 and az.vN > az.v0:
        frasi.append("Sia per il settore che per l'impresa si osserva un andamento positivo dell'indicatore.")
    return " ".join(frasi)


def _evoluzione_struttura1(az, sett):
    """"Il dato e' in forte miglioramento rispetto allo 0,96 del 2021 e supera l'unita' dal 2022, ..." """
    v = az.var or 0
    if v > FORTE:
        giudizio = "in forte miglioramento"
    elif v > LIEVE:
        giudizio = "in miglioramento"
    elif v < -FORTE:
        giudizio = "in forte peggioramento"
    elif v < -LIEVE:
        giudizio = "in peggioramento"
    else:
        giudizio = "sostanzialmente in linea"
    frase = f"Il dato è {giudizio} rispetto {na(az.v0, '', 'a')} del {az.a0}"
    primo = az.primo_anno_sopra(1)
    if az.v0 <= 1 and primo:
        frase += f" e supera l'unità dal {primo}"
    elif az.sempre_sopra(1):
        frase += f" e resta sopra l'unità {tutto_il_periodo(az, 'in')}"
    if sett.valida:
        cresciuta = 'cresciuta' if sett.vN > sett.v0 else 'scesa'
        if az.vN < sett.vN:
            frase += (f", ma rimane inferiore alla mediana del settore, {cresciuta} "
                      f"{na(sett.v0, '', 'da')} {na(sett.vN, '', 'a')} nel quadriennio")
        else:
            frase += (f" e si colloca sopra la mediana del settore, {cresciuta} "
                      f"{na(sett.v0, '', 'da')} {na(sett.vN, '', 'a')} nel quadriennio")
    return frase


def bullet_struttura2(nome_art, az, sett):
    frasi = [_testa(nome_art, az.valore(2024), sett.valore(2024), '')]
    if not az.valida:
        return frasi[0]
    if az.vN > 1:
        frasi.append("Il capitale permanente, dato da patrimonio netto e passività non correnti, copre "
                     "quindi integralmente le immobilizzazioni.")
    else:
        frasi.append("Il capitale permanente, dato da patrimonio netto e passività non correnti, non "
                     "basta a coprire integralmente le immobilizzazioni.")
    frasi.append(percorso_struttura2(az) + ".")
    if sett.valida and az.picco and az.amax == az.aP:
        s_picco = sett.valore(az.amax)
        var_sett_ultimo = var_pct(s_picco, sett.vN)
        if s_picco is not None and sett._sale_fino_a(az.amax) and var_sett_ultimo is not None \
                and var_sett_ultimo > -LIEVE / 5:
            consolida = ("consolida il dato dell'esercizio precedente" if abs(var_sett_ultimo) < 1
                         else "continua a crescere")
            frasi.append(f"Il trend registrato segue quello di settore fino al {az.amax} (il valore mediano "
                         f"dell'indice cresce {na(sett.v0, '', 'da')} {na(s_picco, '', 'a')}) e se ne "
                         f"discosta nel {az.aN}, anno in cui il settore {consolida} con un valore "
                         f"{na(sett.vN, '', 'di')}.")
            return " ".join(frasi)
    if sett.valida:
        frasi.append(f"Nello stesso periodo la mediana di settore {movimento_settore(sett)}.")
    return " ".join(frasi)


def percorso_struttura2(az):
    """"L'indicatore e' passato da 2,22 nel 2021 a 3,37 nel 2023, con una lieve riduzione ..." """
    if az.picco:
        calo = var_pct(az.vmax, az.vN) or 0
        tipo = "una lieve riduzione" if abs(calo) < LIEVE else "una riduzione"
        frase = (f"L'indicatore è passato da {n(az.v0)} nel {az.a0} a {n(az.vmax)} nel {az.amax}, "
                 f"con {tipo} a {n(az.vN)} nel {az.aN}")
    else:
        frase = f"L'indicatore è passato da {n(az.v0)} nel {az.a0} a {n(az.vN)} nel {az.aN}"
    if az.sempre_sopra(1):
        frase += "; resta comunque stabilmente superiore all'unità"
    elif az.vN <= 1:
        frase += "; resta al di sotto dell'unità"
    return frase


def percorso_gearing(az, sett):
    """"L'indicatore aziendale e' diminuito rispetto al 200,14% del 2021 e al 179,01% del 2023, ..." """
    u = '%'
    if not az.valida:
        return ''
    if az.vN < az.v0 and az.vN < az.vP and az.aP != az.a0:
        frase = (f"L'indicatore aziendale è diminuito rispetto {na(az.v0, u, 'a')} del {az.a0} e "
                 f"{na(az.vP, u, 'a')} del {az.aP}")
    elif az.vN < az.v0:
        frase = (f"L'indicatore aziendale è diminuito rispetto {na(az.v0, u, 'a')} del {az.a0}, pur "
                 f"risalendo rispetto {na(az.vP, u, 'a')} del {az.aP}")
    elif az.vN > az.v0:
        frase = f"L'indicatore aziendale è aumentato rispetto {na(az.v0, u, 'a')} del {az.a0}"
    else:
        frase = f"L'indicatore aziendale resta sui livelli del {az.a0}"
    if sett.valida:
        if az.vN > sett.vN * 1.5:
            frase += ", ma resta su un livello significativamente superiore al benchmark"
        elif az.vN > sett.vN:
            frase += ", ma resta superiore al benchmark"
        else:
            # dopo l'inciso "pur risalendo ..." la virgola lo chiude
            frase += ("," if "pur risalendo" in frase else "") + " e si colloca al di sotto del benchmark"
    return frase


def bullet_gearing(nome_art, az, sett):
    u = '%'
    frasi = [_testa(nome_art, az.valore(2024), sett.valore(2024), u)]
    if not az.valida:
        return frasi[0]
    frasi.append(percorso_gearing(az, sett) + ".")
    if sett.valida:
        verbo = "è scesa" if sett.vN < sett.v0 else "è salita"
        stessa_direzione = (sett.vN < sett.v0) == (az.vN < az.v0)
        apertura = "Analogamente, la" if stessa_direzione else "Nello stesso periodo la"
        frase = f"{apertura} mediana settoriale {verbo} {na(sett.v0, u, 'da')} {na(sett.vN, u, 'a')}"
        if az.vN > sett.vN * 1.5:
            frase += ", per cui il divario relativo rispetto al comparto rimane rilevante"
        frasi.append(frase + ".")
    return " ".join(frasi)


def percorso_liquidita(az, chiave):
    """Current e Quick Ratio: soglia dell'unita' e percorso."""
    if not az.valida:
        return ''
    if chiave == 'cr':
        if az.sempre_sopra(1):
            testa = f"L'indicatore è rimasto superiore all'unità {tutto_il_periodo(az, 'per')}, passando"
        else:
            testa = "L'indicatore passa"
        if az.picco:
            return (f"{testa} da {n(az.v0)} nel {az.a0} a {n(az.vmax)} nel {az.amax} e scendendo a "
                    f"{n(az.vN)} nel {az.aN}" if testa.endswith('passando') else
                    f"{testa} da {n(az.v0)} nel {az.a0} a {n(az.vmax)} nel {az.amax}, per poi scendere a "
                    f"{n(az.vN)} nel {az.aN}")
        return f"{testa} da {n(az.v0)} nel {az.a0} a {n(az.vN)} nel {az.aN}"
    # Quick Ratio
    if az.picco:
        return f"Dopo il massimo di {n(az.vmax)} nel {az.amax}, l'indicatore scende a {n(az.vN)} nel {az.aN}"
    if az.minimo:
        return f"Dopo il minimo di {n(az.vmin)} nel {az.amin}, l'indicatore risale a {n(az.vN)} nel {az.aN}"
    verbo = 'sale' if az.vN > az.v0 else 'scende'
    return f"L'indicatore {verbo} da {n(az.v0)} nel {az.a0} a {n(az.vN)} nel {az.aN}"


def bullet_current_ratio(nome_art, az, sett):
    frasi = [_testa(nome_art, az.valore(2024), sett.valore(2024), '')]
    if not az.valida:
        return frasi[0]
    frasi.append(percorso_liquidita(az, 'cr') + ".")
    lettura = lettura_calo_current_ratio(az)
    if lettura:
        frasi.append(lettura)
    if sett.valida:
        frasi.append(f"Nello stesso periodo il dato di settore passa {na(sett.v0, '', 'da')} {na(sett.vN, '', 'a')}.")
    return " ".join(frasi)


def lettura_calo_current_ratio(az):
    if az.picco and az.vN > 1:
        vicino_inizio = abs(var_pct(az.v0, az.vN) or 100) < 15
        if vicino_inizio:
            return (f"La riduzione del {az.aN} non elimina la copertura delle passività correnti, ma "
                    f"segnala un ritorno verso livelli più vicini a quelli di inizio periodo.")
        return f"La riduzione del {az.aN} non elimina la copertura delle passività correnti."
    if az.vN <= 1:
        return "Le attività correnti non bastano a coprire le passività di pari scadenza."
    return ""


def bullet_quick_ratio(nome_art, az, sett):
    frasi = [_testa(nome_art, az.valore(2024), sett.valore(2024), '')]
    if not az.valida:
        return frasi[0]
    sopra, sotto = sopra_in_tutti(az, sett) if sett.valida else ([], [1])
    if az.sempre_sopra(1):
        f = "Il valore è superiore all'unità in tutti gli anni osservati"
        if not sotto:
            f += " e sempre più alto del benchmark settoriale"
        frasi.append(f + ".")
    percorso = percorso_liquidita(az, 'qr')
    if az.vN > 1:
        percorso += (", mantenendo comunque un margine di copertura delle passività correnti con le "
                     "attività prontamente liquidabili senza necessità di smobilizzare le rimanenze di "
                     "magazzino")
    else:
        percorso += ", sotto l'unità: la copertura dipende dallo smobilizzo delle rimanenze"
    frasi.append(percorso + ".")
    if sett.valida:
        if sett.sale_sempre:
            frasi.append(f"Per il settore si registra un progressivo incremento dell'indice che passa "
                         f"da {n(sett.v0)} a {n(sett.vN)} nel periodo considerato.")
        else:
            frasi.append(f"Per il settore l'indice passa da {n(sett.v0)} a {n(sett.vN)} nel periodo considerato.")
    return " ".join(frasi)


def confronto_anni_benchmark(az, sett):
    """"superiore al benchmark in tutti gli anni", con le eccezioni dette per nome."""
    sopra, sotto = sopra_in_tutti(az, sett)
    if not sotto:
        return "è superiore al benchmark in tutti gli anni"
    if len(sotto) == 1:
        a = sotto[0]
        m = sett.valore(a)
        v = az.valore(a)
        if m and abs(v - m) / abs(m) * 100 < 2:
            return (f"è superiore al benchmark in tutti gli anni tranne il {a}, quando si allinea "
                    f"alla mediana di settore")
        return f"è superiore al benchmark in tutti gli anni tranne il {a}"
    if not sopra:
        return "è inferiore al benchmark in tutti gli anni"
    return f"è superiore al benchmark in {len(sopra)} anni su {len(az.punti)}"


def salto_ultimo_anno(az):
    """"...nel 2024 registra un incremento molto marcato rispetto all'1,98 del 2023"."""
    v = az.var_ultimo
    if v is None:
        return ''
    if v > MARCATO:
        return f"nel {az.aN} registra un incremento molto marcato rispetto {na(az.vP, '', 'a')} del {az.aP}"
    if v > 15:
        return f"nel {az.aN} registra un incremento rispetto {na(az.vP, '', 'a')} del {az.aP}"
    if v < -MARCATO:
        return f"nel {az.aN} registra una riduzione molto marcata rispetto {na(az.vP, '', 'a')} del {az.aP}"
    if v < -15:
        return f"nel {az.aN} registra una riduzione rispetto {na(az.vP, '', 'a')} del {az.aP}"
    return f"nel {az.aN} resta vicino al valore del {az.aP}"


def bullet_rotazione(nome_art, az, sett):
    frasi = [_testa(nome_art, az.valore(2024), sett.valore(2024), '')]
    if not az.valida:
        return frasi[0]
    if sett.valida:
        frasi.append(f"Il valore {confronto_anni_benchmark(az, sett)}, ma {salto_ultimo_anno(az)}.")
        s_prec = sett.valore(az.aP)
        frasi.append(f"Il dato mediano di settore, partendo da un valore iniziale di {n(sett.v0)}, si porta "
                     f"a {n(s_prec)} nel {az.aP} fino ad arrivare a {n(sett.vN)} nel {sett.aN}.")
    else:
        frasi.append(f"Il valore {salto_ultimo_anno(az)}.")
    return " ".join(frasi)


# ----------------------------------------------------- slide andamento ----
# Nella presentazione il commento del grafico di andamento e' il solo percorso,
# con la riga di lettura: il confronto con la mediana sta nella slide del 2024.

def percorso_rotazione_ppt(az):
    """"L'indice sale a 2,58 nel 2022, si contrae a 1,98 l'anno seguente e ..." """
    if not az.valida or len(az.punti) < 3:
        return ''
    pezzi = []
    for i, ((a0, v0), (a1, v1)) in enumerate(zip(az.punti, az.punti[1:])):
        var = var_pct(v0, v1) or 0
        if i == 0:
            verbo = 'sale' if v1 > v0 else 'scende'
            pezzi.append(f"L'indice {verbo} {na(v1, '', 'a')} nel {a1}")
        elif i == len(az.punti) - 2:
            if var > MARCATO:
                pezzi.append(f"registra un marcato incremento nel {a1}, fino {na(v1, '', 'a')}")
            elif var < -MARCATO:
                pezzi.append(f"registra una marcata riduzione nel {a1}, fino {na(v1, '', 'a')}")
            else:
                verbo = 'sale' if v1 > v0 else 'scende'
                pezzi.append(f"{verbo} {na(v1, '', 'a')} nel {a1}")
        else:
            verbo = 'si contrae' if v1 < v0 else 'cresce'
            pezzi.append(f"{verbo} {na(v1, '', 'a')} l'anno seguente")
    return ", ".join(pezzi[:-1]) + " e " + pezzi[-1] + "."


def percorso_qr_ppt(az):
    if az.picco:
        return (f"Partendo dal valore di {n(az.v0)} nel {az.a0} e dopo aver raggiunto il picco di "
                f"{n(az.vmax)} nel {az.amax}, l'indicatore scende a {n(az.vN)} nel {az.aN}"
                + (", mantenendo comunque un margine di copertura delle passività correnti con le "
                   "attività prontamente liquidabili senza necessità di smobilizzare le rimanenze di "
                   "magazzino." if az.vN > 1 else ", sotto l'unità."))
    return percorso_liquidita(az, 'qr') + "."


def percorso_struttura1_ppt(az):
    v = az.var or 0
    giudizio = ("in forte miglioramento" if v > FORTE else "in miglioramento" if v > LIEVE else
                "in forte peggioramento" if v < -FORTE else "in peggioramento" if v < -LIEVE else
                "sostanzialmente stabile")
    frase = f"Il ratio è {giudizio} rispetto {na(az.v0, '', 'a')} del {az.a0}"
    primo = az.primo_anno_sopra(1)
    if az.v0 <= 1 and primo:
        frase += f" e supera l'unità dal {primo}"
    return frase + f", attestandosi a {n(az.vN)} nel {az.aN}."


def commento_andamento_ppt(chiave, az, sett):
    """Il testo della card nella slide di andamento di un indicatore."""
    if not az.valida:
        return "Serie storica insufficiente per leggere un andamento."
    if chiave in ('ebitda', 'ebit', 'profitto'):
        testo = percorso_margine(az, chiave) + "."
        if chiave == 'ebitda' and az.picco and az.amax == az.aP:
            gap_prima = (sett.valore(az.aP) - az.vP) if sett.valida and sett.valore(az.aP) is not None else None
            gap_ora = (sett.vN - az.vN) if sett.valida else None
            coda = ''
            if gap_prima is not None and gap_ora is not None and gap_ora > 0:
                coda = (" e amplia il divario rispetto al comparto" if gap_ora > gap_prima
                        else " pur riducendo il divario rispetto al comparto")
            testo += f" Il {az.aN} interrompe quindi il miglioramento osservato fino al {az.amax}{coda}."
        elif chiave == 'ebit' and az.vN < az.vP:
            precedenti = [v for a, v in az.punti if az.aN - 2 <= a < az.aN]
            testo += f" La contrazione di {pp(az.vP - az.vN)} rispetto al {az.aP}"
            if len(precedenti) == 2 and all(az.vN < v for v in precedenti):
                testo += " riporta il margine al di sotto dei livelli osservati nei due anni precedenti"
            testo += "."
        elif chiave == 'profitto' and az.picco and az.amax == az.aP:
            testo += (f" La dinamica conferma che il recupero di redditività registrato fino al "
                      f"{az.amax} non si è consolidato nell'ultimo esercizio.")
        return testo
    if chiave == 'strut1':
        return percorso_struttura1_ppt(az)
    if chiave == 'strut2':
        return percorso_struttura2(az) + "."
    if chiave == 'gearing':
        return percorso_gearing(az, sett) + "."
    if chiave == 'cr':
        lettura = lettura_calo_current_ratio(az)
        return percorso_liquidita(az, 'cr') + "." + (" " + lettura if lettura else "")
    if chiave == 'qr':
        return percorso_qr_ppt(az)
    if chiave == 'rotazione':
        return percorso_rotazione_ppt(az)
    return ""


# --------------------------------------------------------- dispatcher ----

NOMI_ART = {
    'ebitda': 'Il Margine EBITDA', 'ebit': 'Il Margine EBIT', 'profitto': 'Il Margine di Profitto',
    'strut1': "L'Indice di Struttura di 1° livello", 'strut2': "L'Indice di Struttura di 2° livello",
    'gearing': 'Il Gearing', 'cr': 'Il Current Ratio', 'qr': 'Il Quick Ratio',
    'rotazione': "L'Indice di Rotazione del Capitale Investito",
}


def bullet_indicatore(chiave, serie_az, serie_set):
    """Il bullet del capitolo di posizionamento per un indicatore."""
    az, sett = Serie(serie_az), Serie(serie_set)
    nome = NOMI_ART[chiave]
    if chiave in ('ebitda', 'ebit', 'profitto'):
        return bullet_margine(chiave, nome, az, sett)
    if chiave == 'strut1':
        return bullet_struttura1(nome, az, sett)
    if chiave == 'strut2':
        return bullet_struttura2(nome, az, sett)
    if chiave == 'gearing':
        return bullet_gearing(nome, az, sett)
    if chiave == 'cr':
        return bullet_current_ratio(nome, az, sett)
    if chiave == 'qr':
        return bullet_quick_ratio(nome, az, sett)
    return bullet_rotazione(nome, az, sett)


# ======================================================================
# Paragrafi dell'Analisi degli Equilibri e della sintesi
# ======================================================================
# `dati` e' un dizionario con:
#   nome                 ragione sociale
#   az, sett             {chiave: {anno: valore}} per i nove indicatori
#   rating_eco/_patr/_fin/_tot
#   comp_az, comp_sett   incidenze 2024 sul Valore della Produzione, in %:
#                        'venduto', 'oneri_gestione', 'finanziari'

def _serie(dati, chiave):
    return Serie(dati['az'].get(chiave, {})), Serie(dati['sett'].get(chiave, {}))


def _sotto_mediana(az, sett, inverso=False):
    if not (az.valida and sett.valida):
        return False
    return az.vN > sett.vN if inverso else az.vN < sett.vN


def _elenco(voci):
    """"A, B e C"; "ed" davanti a una parola che inizia per e ("EBITDA ed EBIT")."""
    voci = [v for v in voci if v]
    if len(voci) <= 1:
        return ''.join(voci)
    congiunzione = " ed " if voci[-1][:1].lower() == 'e' else " e "
    return ", ".join(voci[:-1]) + congiunzione + voci[-1]


# --- Executive Summary --------------------------------------------------

def descr_rating_tot(dati):
    """Dopo 'attribuisce a X un Rating Combinato "CBA".' """
    eb, se = _serie(dati, 'ebitda')
    s1, _ = _serie(dati, 'strut1')
    s2, _ = _serie(dati, 'strut2')
    cr, _ = _serie(dati, 'cr')
    qr, _ = _serie(dati, 'qr')
    margini = [_serie(dati, k) for k in ('ebitda', 'ebit', 'profitto')]

    pezzi = []
    if cr.valida and qr.valida and cr.vN > 1 and qr.vN > 1:
        pezzi.append("la liquidità di breve periodo è adeguata")
    elif cr.valida and cr.vN > 1:
        pezzi.append("la liquidità di breve periodo è adeguata solo contando le rimanenze")
    else:
        pezzi.append("la liquidità di breve periodo non copre le passività correnti")
    if s1.valida and s2.valida and s1.vN > 1 and s2.vN > 1:
        pezzi.append("la struttura delle fonti copre le immobilizzazioni con capitale permanente")
    elif s2.valida and s2.vN > 1:
        pezzi.append("la struttura delle fonti copre le immobilizzazioni solo con il contributo dei debiti a "
                     "medio-lungo termine")
    else:
        pezzi.append("la struttura delle fonti non copre integralmente le immobilizzazioni")

    positivi = all(a.valida and a.vN > 0 for a, _ in margini)
    sotto = [a.valida and s.valida and a.vN < s.vN for a, s in margini]
    if positivi and all(sotto):
        red = ("la redditività appare soddisfacente, sebbene dal confronto con il comparto di "
               "riferimento si collochi per tutti e tre i margini al di sotto della mediana di settore")
    elif positivi and not any(sotto):
        red = "la redditività si colloca per tutti e tre i margini al di sopra della mediana di settore"
    elif positivi:
        red = "la redditività è positiva, con un confronto con il comparto che varia da margine a margine"
    else:
        red = "la redditività è negativa su almeno uno dei tre margini"
    articolata = len({dati['rating_eco'], dati['rating_patr'], dati['rating_fin']}) > 1
    apertura = ("Questo risultato evidenzia una situazione articolata" if articolata
                else "Questo risultato evidenzia un profilo omogeneo")
    return f"{apertura}: {pezzi[0]}, {pezzi[1]} e {red}."


def descr_rating_eco(dati, testo_attuale_coda):
    """
    "nel 2024 i tre margini analizzati risultano inferiori alle rispettive mediane
    settoriali: il divario e' particolarmente evidente per i margini EBITDA ed EBIT e
    si accompagna a una contrazione rispetto al 2023. <coda>"
    """
    nomi = {'ebitda': 'EBITDA', 'ebit': 'EBIT', 'profitto': 'di Profitto'}
    margini = {k: _serie(dati, k) for k in nomi}
    sotto = [k for k, (a, s) in margini.items() if _sotto_mediana(a, s)]
    if len(sotto) == 3:
        testa = "nel 2024 i tre margini analizzati risultano inferiori alle rispettive mediane settoriali"
    elif sotto:
        testa = ("nel 2024 risultano sotto la mediana di settore i margini "
                 + _elenco([nomi[k] for k in sotto]))
    else:
        testa = "nel 2024 i tre margini si collocano sopra le rispettive mediane settoriali"
    if sotto:
        # divario relativo: quanto manca in proporzione alla mediana
        relativi = {k: (margini[k][1].vN - margini[k][0].vN) / abs(margini[k][1].vN) * 100
                    for k in sotto if abs(margini[k][1].vN) > 1e-9}
        evidenti = [k for k in ('ebitda', 'ebit', 'profitto') if relativi.get(k, 0) >= FORTE]
        eb = margini['ebitda'][0]
        contrazione = eb.valida and eb.vN < eb.vP
        coda = ''
        if evidenti and len(evidenti) < len(sotto):
            parola = 'il margine' if len(evidenti) == 1 else 'i margini'
            coda = (f": il divario \u00e8 particolarmente evidente per {parola} "
                    f"{_elenco([nomi[k] for k in evidenti])}")
            if contrazione:
                coda += f" e si accompagna a una contrazione rispetto al {eb.aP}"
        elif contrazione:
            coda = f": il divario si accompagna a una contrazione rispetto al {eb.aP}"
        testa += coda
    return f"{testa}. {testo_attuale_coda}".strip()


def intro_margini_coda(dati):
    """Frase aggiunta dopo le cifre del Benchmark Economico."""
    margini = [_serie(dati, k)[0] for k in ('ebitda', 'ebit', 'profitto')]
    if all(a.valida and a.picco and a.amax == a.aP for a in margini):
        return (f"Il dato del {margini[0].aN} va inoltre letto alla luce della dinamica precedente, che "
                f"aveva visto un miglioramento dei margini nel {margini[0].aP} seguito da una contrazione "
                f"nell'ultimo esercizio.")
    return ''


def sintesi_rating_combinato(dati):
    """Paragrafo dopo la Tabella 5: cosa riassume il Rating Combinato."""
    r = dati
    testo = (f"riassume i risultati dei tre benchmark: classe {r['rating_eco']} per l'Equilibrio "
             f"Economico, classe {r['rating_patr']} per l'Equilibrio Patrimoniale e classe "
             f"{r['rating_fin']} per l'Equilibrio Finanziario.")
    # quanti dei nove indicatori stanno dal lato sfavorevole della mediana
    inversi = {'gearing'}
    sfavorevoli = 0
    for k in ('ebitda', 'ebit', 'profitto', 'strut1', 'strut2', 'gearing', 'cr', 'qr', 'rotazione'):
        a, s = _serie(dati, k)
        if _sotto_mediana(a, s, k in inversi):
            sfavorevoli += 1
    ambiti = []
    margini = [_serie(dati, k) for k in ('ebitda', 'ebit', 'profitto')]
    if all(_sotto_mediana(a, s) for a, s in margini):
        ambiti.append("della redditività")
    g, gs = _serie(dati, 'gearing')
    if g.valida and gs.valida and g.vN > gs.vN * 1.5:
        ambiti.append("della situazione di indebitamento")
    if sfavorevoli >= 6:
        giudizio = "indicatori in linea generale inferiori ai riferimenti di settore"
    elif sfavorevoli >= 4:
        giudizio = "indicatori in parte inferiori ai riferimenti di settore"
    else:
        giudizio = "indicatori in prevalenza allineati o superiori ai riferimenti di settore"
    testo += f" La sintesi evidenzia {giudizio}"
    if ambiti:
        testo += (", con divari particolarmente evidenti per quanto attiene alla sfera "
                  + _elenco(ambiti))
    return testo + "."


def sintesi_sul_piano_economico(dati):
    margini = [_serie(dati, k) for k in ('ebitda', 'ebit', 'profitto')]
    ca, cs = dati.get('comp_az', {}), dati.get('comp_sett', {})
    op_az = ca.get('venduto', 0) + ca.get('oneri_gestione', 0)
    op_set = cs.get('venduto', 0) + cs.get('oneri_gestione', 0)
    if all(_sotto_mediana(a, s) for a, s in margini):
        testo = ("Sul piano economico, il confronto con il settore mostra un divario nei tre margini che "
                 "può essere spiegato attraverso l'analisi dell'incidenza delle diverse componenti di "
                 "costo rispetto al Valore della Produzione")
        if op_az > op_set:
            testo += (", dalla quale emerge una maggiore incidenza dei costi operativi rispetto alla "
                      "mediana settoriale")
        return testo + "."
    if not any(_sotto_mediana(a, s) for a, s in margini):
        return ("Sul piano economico, i tre margini si collocano sopra le rispettive mediane di settore: "
                "la gestione caratteristica trasforma il Valore della Produzione in reddito operativo "
                "meglio del comparto.")
    return ("Sul piano economico, il confronto con il settore è differenziato da margine a margine e va "
            "letto insieme all'incidenza delle diverse componenti di costo sul Valore della Produzione.")


def sintesi_sul_piano_patrimoniale(dati):
    s1, _ = _serie(dati, 'strut1')
    s2, _ = _serie(dati, 'strut2')
    g, gs = _serie(dati, 'gearing')
    frasi = []
    if s1.valida and s2.valida and s1.vN > 1 and s2.vN > 1:
        frasi.append("Sul piano patrimoniale, gli indici di struttura risultano superiori all'unità e "
                     "confermano la copertura delle immobilizzazioni mediante fonti durevoli.")
    else:
        frasi.append("Sul piano patrimoniale, gli indici di struttura non garantiscono la piena copertura "
                     "delle immobilizzazioni mediante fonti durevoli.")
    if g.valida and gs.valida:
        if g.vN > gs.vN * 1.5:
            f = f"Il Gearing, pari {na(g.vN, '%', 'a')}, è invece molto superiore alla mediana di settore"
            if g.vN < g.vP:
                f += f" e, pur essendo diminuito rispetto al {g.aP}, resta su livelli elevati"
            frasi.append(f + ".")
            frasi.append("Il monitoraggio dell'indebitamento e del relativo costo assume pertanto "
                         "particolare rilievo.")
        elif g.vN > gs.vN:
            frasi.append(f"Il Gearing, pari {na(g.vN, '%', 'a')}, è superiore alla mediana di settore.")
        else:
            frasi.append(f"Il Gearing, pari {na(g.vN, '%', 'a')}, è contenuto rispetto alla mediana di "
                         f"settore.")
    return " ".join(frasi)


def sintesi_sul_piano_finanziario(dati):
    cr, _ = _serie(dati, 'cr')
    qr, _ = _serie(dati, 'qr')
    ro, ros = _serie(dati, 'rotazione')
    frasi = []
    if cr.valida and qr.valida and cr.vN > 1 and qr.vN > 1:
        frasi.append("Sul piano finanziario, Current Ratio e Quick Ratio restano superiori all'unità ed "
                     "evidenziano la capacità dell'impresa di coprire le passività correnti anche senza "
                     "contare sulle rimanenze.")
    elif cr.valida and cr.vN > 1:
        frasi.append("Sul piano finanziario, il Current Ratio resta superiore all'unità ma il Quick Ratio "
                     "no: la copertura delle passività correnti dipende dalle rimanenze.")
    else:
        frasi.append("Sul piano finanziario, le attività correnti non bastano a coprire le passività "
                     "di pari scadenza.")
    if ro.valida and ros.valida:
        f = f"L'Indice di Rotazione del Capitale Investito raggiunge {n(ro.vN)} nel {ro.aN}"
        if ro.vN > ros.vN * 2:
            f += f", valore eccezionalmente superiore alla mediana settoriale di {n(ros.vN)}"
        elif ro.vN > ros.vN:
            f += f", sopra la mediana settoriale di {n(ros.vN)}"
        else:
            f += f", sotto la mediana settoriale di {n(ros.vN)}"
        if (ro.var_ultimo or 0) > MARCATO:
            f += (f"; la dinamica, caratterizzata da un forte incremento rispetto al {ro.aP}, merita di "
                  f"essere interpretata considerando contestualmente l'evoluzione del capitale investito e "
                  f"del Valore della Produzione")
        frasi.append(f + ".")
    return " ".join(frasi)


def sintesi_finale(dati):
    s1, _ = _serie(dati, 'strut1')
    s2, _ = _serie(dati, 'strut2')
    g, gs = _serie(dati, 'gearing')
    cr, _ = _serie(dati, 'cr')
    qr, _ = _serie(dati, 'qr')
    margini = {k: _serie(dati, k)[0] for k in ('ebitda', 'ebit', 'profitto')}

    tenuta = cr.valida and qr.valida and cr.vN > 1 and qr.vN > 1
    copertura = s1.valida and s2.valida and s1.vN > 1 and s2.vN > 1
    debito_alto = g.valida and gs.valida and g.vN > gs.vN
    parti = []
    parti.append("una buona tenuta finanziaria" if tenuta else "una tenuta finanziaria da presidiare")
    struttura = ("una struttura patrimoniale in grado di coprire gli investimenti mediante fonti durevoli"
                 if copertura else "una struttura patrimoniale che non copre integralmente gli investimenti")
    frase = f"In sintesi, il profilo {g.aN if g.valida else 2024} combina {parti[0]} con {struttura}"
    if debito_alto:
        frase += ", ma con un ricorso al debito ancora elevato"
    frasi = [frase + "."]
    calati = [k for k, a in margini.items() if a.valida and a.vN < a.vP]
    nomi = {'ebitda': 'EBITDA', 'ebit': 'EBIT', 'profitto': 'Margine di Profitto'}
    if len(calati) == 3:
        frasi.append(f"La principale discontinuità rispetto al {margini['ebitda'].aP} riguarda la "
                     f"redditività, con la riduzione di EBITDA, EBIT e Margine di Profitto.")
    elif calati:
        frasi.append(f"Rispetto al {margini['ebitda'].aP} si riduce {_elenco([nomi[k] for k in calati])}.")
    obiettivi = []
    if any(_sotto_mediana(*_serie(dati, k)) for k in ('ebitda', 'ebit', 'profitto')):
        obiettivi.append("sul recupero dei margini")
    if debito_alto:
        obiettivi.append("sul controllo dell'indebitamento")
    if obiettivi:
        coda = ", mantenendo il presidio della liquidità" if tenuta else ", insieme al presidio della liquidità"
        frasi.append(f"L'attenzione gestionale può quindi concentrarsi {_elenco(obiettivi)}{coda}.")
    return " ".join(frasi)


# --- Analisi degli Equilibri: area economica ------------------------------

def _percorso_ebitda_breve(az, u='%'):
    """"l'EBITDA aziendale passa dal 5,84% del 2021 all'8,14% del 2023, per poi ridursi al 5,43% nel 2024" """
    if az.picco:
        return (f"passa {na(az.v0, u, 'da')} del {az.a0} {na(az.vmax, u, 'a')} del {az.amax}, per poi "
                f"ridursi {na(az.vN, u, 'a')} nel {az.aN}")
    verbo = 'sale' if az.vN > az.v0 else 'scende'
    return f"{verbo} {na(az.v0, u, 'da')} del {az.a0} {na(az.vN, u, 'a')} del {az.aN}"


def ebitda_differenza(dati):
    """Coda del paragrafo sul Margine EBITDA (Tabella 8 e Figura 3)."""
    az, sett = _serie(dati, 'ebitda')
    if not (az.valida and sett.valida) or az.vN >= sett.vN:
        return ''
    ca, cs = dati.get('comp_az', {}), dati.get('comp_sett', {})
    op_az = ca.get('venduto', 0) + ca.get('oneri_gestione', 0)
    op_set = cs.get('venduto', 0) + cs.get('oneri_gestione', 0)
    if op_az > op_set:
        return (f"La differenza di {pp(sett.vN - az.vN)} riscontrata per l'ultimo anno è coerente con "
                f"una maggiore incidenza delle componenti operative sul Valore della Produzione (Tabella 7).")
    return (f"La differenza di {pp(sett.vN - az.vN)} riscontrata per l'ultimo anno va letta insieme alla "
            f"composizione del Valore della Produzione (Tabella 7).")


def ebitda_composizione(dati):
    az, sett = _serie(dati, 'ebitda')
    ca, cs = dati.get('comp_az', {}), dati.get('comp_sett', {})
    if not ca or not cs:
        return ''
    divario = ("il divario" if az.valida and sett.valida and az.vN < sett.vN else "il differenziale")
    return (f"Nel {az.aN if az.valida else 2024} {divario} dell'EBITDA rispetto al settore è "
            f"riconducibile soprattutto alla diversa composizione del Valore della Produzione. Ponendo "
            f"quest'ultimo pari a 100, {dati['nome']} presenta un'incidenza del costo del venduto pari "
            f"{na(ca.get('venduto'), '%', 'a')}, contro {na(cs.get('venduto'), '%')} della mediana "
            f"settoriale, mentre gli oneri di gestione incidono per {na(ca.get('oneri_gestione'), '%')}, "
            f"contro {na(cs.get('oneri_gestione'), '%')} del settore.")


def ebitda_trend(dati):
    az, sett = _serie(dati, 'ebitda')
    ca, cs = dati.get('comp_az', {}), dati.get('comp_sett', {})
    frasi = []
    if ca and cs:
        una_sopra_una_sotto = ((ca.get('venduto', 0) > cs.get('venduto', 0))
                               != (ca.get('oneri_gestione', 0) > cs.get('oneri_gestione', 0)))
        if una_sopra_una_sotto:
            frasi.append("La diversa struttura dei costi suggerisce quindi di leggere il minore EBITDA non "
                         "come effetto di una singola voce, ma come risultato della combinazione delle "
                         "diverse componenti della gestione caratteristica.")
        else:
            frasi.append("L'incidenza delle due voci va nella stessa direzione e spiega da sola buona "
                         "parte del differenziale di EBITDA rispetto al settore.")
    if az.valida and sett.valida:
        meno = (az.var or 0) < (sett.var or 0)
        apertura = ("Anche il confronto temporale evidenzia una dinamica meno favorevole rispetto al "
                    "comparto" if meno else
                    "Il confronto temporale evidenzia invece una dinamica più favorevole rispetto al comparto")
        verbo = 'cresce' if sett.vN > sett.v0 else 'scende'
        frase = (f"{apertura}: l'EBITDA aziendale {_percorso_ebitda_breve(az)}; nello stesso periodo la "
                 f"mediana settoriale {verbo} {na(sett.v0, '%', 'da')} {na(sett.vN, '%', 'a')}")
        gap0 = sett.v0 - az.v0
        gapN = sett.vN - az.vN
        if gapN > gap0 > -1e9 and gapN > 0:
            frase += (", con un miglioramento complessivo che rende più ampio nel "
                      f"{az.aN} il differenziale relativo della società" if sett.vN > sett.v0 else
                      f", con un differenziale che nel {az.aN} si amplia")
        frasi.append(frase + " (Tabelle 7, 8 e Figura 4).")
    return " ".join(frasi)


def ebit_paragrafo(dati):
    az, sett = _serie(dati, 'ebit')
    if not az.valida:
        return ''
    frasi = [f"Analogamente, il Margine EBIT dell'azienda si attesta {na(az.vN, '%', 'a')} nell'ultimo "
             f"esercizio" + (f" contro {na(sett.vN, '%')} del settore" if sett.valida else "") + "."]
    if az.picco and az.amax == az.aP:
        frasi.append(f"Dopo il miglioramento registrato tra il {az.a0} ({n(az.v0, '%')}) e il {az.amax} "
                     f"({n(az.vmax, '%')}), il margine si riduce di {pp(az.vP - az.vN)} nel {az.aN}.")
    elif az.valida:
        if az.vN == az.vP:
            frasi.append(f"Nell'ultimo esercizio il margine resta invariato rispetto al {az.aP}.")
        else:
            verbo = 'cresce' if az.vN > az.vP else 'si riduce'
            frasi.append(f"Nell'ultimo esercizio il margine {verbo} di {pp(az.vN - az.vP)} rispetto al {az.aP}.")
    if sett.valida:
        sett_sale_ultimo = sett.vN > sett.vP
        az_scende_ultimo = az.vN < az.vP
        if az_scende_ultimo and sett_sale_ultimo:
            frasi.append(f"La dinamica aziendale si discosta quindi da quella del settore: la mediana passa "
                         f"{na(sett.v0, '%', 'da')} nel {sett.a0} {na(sett.vN, '%', 'a')} nel {sett.aN}, con un "
                         f"incremento anche nell'ultimo esercizio, a differenza della società che invece "
                         f"registra una contrazione.")
            frasi.append(f"Il confronto evidenzia pertanto che il ridimensionamento della redditività "
                         f"operativa nel {az.aN} non è un fenomeno condiviso dal benchmark settoriale")
        elif az_scende_ultimo:
            frasi.append(f"Anche la mediana di settore si riduce nell'ultimo esercizio, passando "
                         f"{na(sett.vP, '%', 'da')} {na(sett.vN, '%', 'a')}: il ridimensionamento della "
                         f"redditività operativa riflette in parte una dinamica comune al comparto")
        else:
            frasi.append(f"Nel quadriennio la mediana di settore passa {na(sett.v0, '%', 'da')} del "
                         f"{sett.a0} {na(sett.vN, '%', 'a')} del {sett.aN}")
    return " ".join(frasi) + " (Tabella 9 e Figura 5)."


def profitto_paragrafo(dati):
    az, sett = _serie(dati, 'profitto')
    if not az.valida:
        return ''
    frasi = [f"Nel {az.aN} il Margine di Profitto è pari {na(az.vN, '%', 'a')}"
             + (f", contro {na(sett.vN, '%')} della mediana settoriale" if sett.valida else "") + "."]
    if az.picco and az.amax == az.aP:
        progressivo = "progressivo " if az._sale_fino_a(az.amax) else ""
        frasi.append(f"Dopo il {progressivo}miglioramento {na(az.v0, '%', 'da')} del {az.a0} "
                     f"{na(az.vmax, '%', 'a')} del {az.amax}, l'indicatore si riduce {na(az.vN, '%', 'a')} "
                     f"nel {az.aN}.")
        precedenti = [v for a, v in az.punti if az.aN - 2 <= a < az.aN]
        f = f"La contrazione di {pp(az.vP - az.vN)} rispetto al {az.aP}"
        if len(precedenti) == 2 and all(az.vN < v for v in precedenti):
            f += " riporta la redditività ante imposte al di sotto dei livelli registrati nei due esercizi precedenti"
        frasi.append(f + ".")
    if sett.valida and az.vN < az.vP:
        calo_sett = var_pct(sett.vP, sett.vN)
        calo_az = var_pct(az.vP, az.vN)
        if calo_sett is not None and calo_sett < 0:
            lieve = "una lieve riduzione" if abs(calo_sett) < LIEVE else "una riduzione"
            f = (f"Anche il settore evidenzia {lieve} nel {sett.aN}, passando dalla mediana "
                 f"{na(sett.vP, '%', 'di')} del {sett.aP} {na(sett.vN, '%', 'a')}")
            if calo_az is not None and calo_az < calo_sett:
                f += ("; tuttavia, la flessione aziendale è più marcata e porta ad ampliare il divario "
                      "rispetto al benchmark")
            frasi.append(f)
        else:
            frasi.append(f"Il settore, al contrario, nel {sett.aN} cresce {na(sett.vP, '%', 'da')} "
                         f"{na(sett.vN, '%', 'a')}")
        return " ".join(frasi) + " (Tabella 10 e Figura 7)."
    return " ".join(frasi).rstrip('.') + " (Tabella 10 e Figura 7)."


def profitto_struttura(dati):
    ca, cs = dati.get('comp_az', {}), dati.get('comp_sett', {})
    if not ca or not cs:
        return ''
    frasi = ["Il confronto tra la società e il settore mostra che il principale divario economico si "
             "concentra nella gestione caratteristica." if (ca.get('venduto', 0) + ca.get('oneri_gestione', 0))
             > (cs.get('venduto', 0) + cs.get('oneri_gestione', 0)) else
             "Il confronto tra la società e il settore mostra una gestione caratteristica meno onerosa "
             "di quella mediana."]
    frasi.append(f"Nel 2024 il costo del venduto incide sul Valore della Produzione per "
                 f"{na(ca.get('venduto'), '%')}, contro {na(cs.get('venduto'), '%')} della mediana, mentre "
                 f"gli oneri di gestione incidono per {na(ca.get('oneri_gestione'), '%')}, contro "
                 f"{na(cs.get('oneri_gestione'), '%')} del settore.")
    fa, fs = ca.get('finanziari'), cs.get('finanziari')
    if _ok(fa) and _ok(fs):
        if fa <= 0 and fs <= 0:
            frasi.append(f"Sul fronte finanziario, l'incidenza degli oneri finanziari aziendali è pari "
                         f"{na(abs(fa), '%', 'a')} del VdP, contro {na(abs(fs), '%')} della mediana.")
            if abs(fa) < abs(fs):
                frasi.append("Ne deriva che, pur in presenza di un'incidenza finanziaria inferiore al benchmark, "
                             "la redditività finale resta più contenuta soprattutto per effetto della "
                             "struttura della gestione caratteristica (Tabella 10 e Figura 8).")
                return " ".join(frasi)
            frasi.append("L'incidenza finanziaria superiore al benchmark si somma quindi al peso della "
                         "gestione caratteristica nel comprimere la redditività finale (Tabella 10 e Figura 8).")
            return " ".join(frasi)
        frasi.append(f"Sul fronte finanziario, la gestione produce un saldo pari {na(fa, '%', 'a')} del VdP, "
                     f"contro {na(fs, '%')} della mediana.")
    return " ".join(frasi) + " (Tabella 10 e Figura 8)."


def eco_conclusione_1(dati):
    margini = {k: _serie(dati, k) for k in ('ebitda', 'ebit', 'profitto')}
    tutti_sotto = all(_sotto_mediana(a, s) for a, s in margini.values())
    tutti_calo = all(a.valida and a.vN < a.vP for a, _ in margini.values())
    a0 = margini['ebitda'][0]
    frasi = [f"Nel {a0.aN} l'Equilibrio Economico appartiene alla classe “{dati['rating_eco']}”."]
    riflette = ("valori di EBITDA, EBIT e Margine di Profitto inferiori alle rispettive mediane settoriali"
                if tutti_sotto else "un confronto con le mediane settoriali differenziato fra i tre margini")
    if tutti_calo:
        riflette += f" e una contrazione dei margini rispetto al {a0.aP}"
    frasi.append(f"Tale posizionamento riflette {riflette}.")
    settore_su = a0.valida and all(s.valida and s._sale_fino_a(a0.aP) and s.vN > s.v0 for _, s in margini.values())
    picco_az = all(a.valida and a.picco and a.amax == a.aP for a, _ in margini.values())
    if settore_su and picco_az:
        frasi.append(f"Il confronto temporale aggiunge un elemento importante: il settore ha progressivamente "
                     f"migliorato i propri valori mediani tra il {a0.a0} e il {a0.aP} e mantenuto, nel "
                     f"{a0.aN}, livelli di redditività superiori a quelli iniziali del quadriennio; la "
                     f"società, invece, dopo il miglioramento fino al {a0.aP}, ha registrato nell'ultimo "
                     f"esercizio una riduzione di tutti e tre i margini.")
        frasi.append(f"Ne deriva che il {a0.aN} rappresenta per l'impresa non solo un arretramento rispetto al "
                     f"proprio massimo recente, ma anche un peggioramento della posizione relativa rispetto al "
                     f"comparto.")
    return " ".join(frasi)


def eco_conclusione_2(dati):
    classi = {'eco': dati['rating_eco'], 'patr': dati['rating_patr'], 'fin': dati['rating_fin']}
    peggiore = max(classi.values())
    ca, cs = dati.get('comp_az', {}), dati.get('comp_sett', {})
    costi_alti = (ca.get('venduto', 0) + ca.get('oneri_gestione', 0)) > (cs.get('venduto', 0) + cs.get('oneri_gestione', 0))
    if classi['eco'] == peggiore and peggiore != 'A':
        testo = "L'area economica rappresenta pertanto il principale ambito di monitoraggio emerso dall'analisi"
        if costi_alti:
            testo += (": il recupero dei margini richiede di comprendere le determinanti della maggiore "
                      "incidenza dei costi operativi e di verificare l'evoluzione della gestione finanziaria")
        return testo + "."
    return ("L'area economica non rappresenta il principale ambito di attenzione emerso dall'analisi, ma "
            "l'evoluzione dei margini resta da seguire nei prossimi esercizi.")


# --- Analisi degli Equilibri: area patrimoniale ---------------------------

def patr_intro_1(dati):
    s1, _ = _serie(dati, 'strut1')
    s2, _ = _serie(dati, 'strut2')
    if s1.valida and s2.valida and s1.vN > s1.v0 and s2.vN > s2.v0:
        return (f"{periodo_anni(s1)} la struttura patrimoniale mostra un miglioramento degli "
                f"indici di copertura delle immobilizzazioni.")
    if s1.valida and s2.valida and s1.vN < s1.v0 and s2.vN < s2.v0:
        return (f"{periodo_anni(s1)} la struttura patrimoniale mostra un indebolimento degli "
                f"indici di copertura delle immobilizzazioni.")
    return (f"Nel quadriennio la struttura patrimoniale mostra un andamento differenziato fra i due indici "
            f"di copertura delle immobilizzazioni.")


def patr_intro_2(dati):
    s1, t1 = _serie(dati, 'strut1')
    s2, t2 = _serie(dati, 'strut2')
    if not (s1.valida and s2.valida):
        return ''
    frasi = [f"L'Indice di Struttura di 1° livello passa da {n(s1.v0)} a {n(s1.vN)} e quello di 2° "
             f"livello da {n(s2.v0)} a {n(s2.vN)}"]
    if s1.vN > 1 and s2.vN > 1:
        frasi[0] += f"; entrambi risultano superiori all'unità nel {s1.aN}."
    else:
        frasi[0] += "."
    if t1.valida and t2.valida:
        frasi.append(f"Il confronto con il trend settoriale è tuttavia differenziato: le mediane del "
                     f"settore {'crescono' if t1.vN > t1.v0 else 'si muovono'} da {n(t1.v0)} a {n(t1.vN)} per "
                     f"il primo indice e da {n(t2.v0)} a {n(t2.vN)} per il secondo.")
        sotto = s1.vN < t1.vN and s2.vN < t2.vN
        migliora = s1.vN > s1.v0 and s2.vN > s2.v0
        if migliora and sotto:
            frasi.append("Pertanto, la società migliora significativamente la propria struttura, ma a un "
                         "ritmo che non le consente di colmare il divario con il comparto.")
        elif migliora:
            frasi.append("Pertanto, la società migliora la propria struttura e si colloca sopra le mediane "
                         "del comparto.")
    return " ".join(frasi)


def struttura1_paragrafo(dati):
    s1, _ = _serie(dati, 'strut1')
    if not s1.valida:
        return ''
    frasi = [f"L'Indice di Struttura di 1° livello è pari a {n(s1.vN)} nel {s1.aN}."]
    if s1.vN > 1:
        frasi.append("Il valore superiore all'unità indica che il patrimonio netto è sufficiente a coprire "
                     "le immobilizzazioni.")
    else:
        frasi.append("Il valore inferiore all'unità indica che il patrimonio netto non basta a coprire le "
                     "immobilizzazioni.")
    v = s1.var or 0
    primo = s1.primo_anno_sopra(1)
    if v > FORTE:
        f = (f"Il dato evidenzia un miglioramento significativo rispetto al {s1.a0}, quando l'indice era "
             f"pari a {n(s1.v0)}")
    elif v > LIEVE:
        f = f"Il dato evidenzia un miglioramento rispetto al {s1.a0}, quando l'indice era pari a {n(s1.v0)}"
    elif v < -LIEVE:
        f = f"Il dato evidenzia un peggioramento rispetto al {s1.a0}, quando l'indice era pari a {n(s1.v0)}"
    else:
        f = f"Il dato è sostanzialmente in linea con il {s1.a0}, quando l'indice era pari a {n(s1.v0)}"
    if s1.v0 <= 1 and primo:
        f += f", e conferma il superamento della soglia di equilibrio a partire dal {primo}"
    return " ".join(frasi) + " " + f + " (Tabella 11 e Figure 9-10)."


def struttura2_paragrafo(dati):
    s2, t2 = _serie(dati, 'strut2')
    if not s2.valida:
        return ''
    frasi = [f"L'Indice di Struttura di 2° livello raggiunge {n(s2.vN)} nel {s2.aN}, indicando che "
             f"patrimonio netto e passività non correnti "
             + ("sono più che sufficienti a garantire la copertura integrale delle immobilizzazioni."
                if s2.vN > 1.5 else "sono sufficienti a coprire le immobilizzazioni." if s2.vN > 1 else
                "non bastano a coprire integralmente le immobilizzazioni.")]
    if s2.vN < s2.vP:
        lieve = "in lieve diminuzione" if abs(var_pct(s2.vP, s2.vN) or 0) < LIEVE else "in diminuzione"
        f = f"Il valore è {lieve} rispetto {na(s2.vP, '', 'a')} del {s2.aP}"
        if s2.vN > 1.5:
            f += ", ma resta ampiamente superiore all'unità"
        frasi.append(f + ".")
    if t2.valida:
        dinamica = "una dinamica crescente" if t2.sale_sempre or t2.vN > t2.v0 else "una dinamica decrescente"
        invece = " invece" if (t2.vN > t2.vP) != (s2.vN > s2.vP) else ""
        frasi.append(f"La mediana settoriale segue{invece} {dinamica} nel quadriennio, da {n(t2.v0)} nel "
                     f"{t2.a0} a {n(t2.vN)} nel {t2.aN}.")
        if s2.vN > s2.v0 and s2.vN < s2.vP and t2.vN >= t2.vP:
            frasi.append(f"La società presenta quindi un miglioramento complessivo rispetto al {s2.a0}, ma "
                         f"nell'ultimo esercizio interrompe parzialmente la propria crescita, mentre il "
                         f"comparto continua a rafforzare la copertura strutturale")
            return " ".join(frasi) + " (Tabella 12 e Figure 11-12)."
    return " ".join(frasi).rstrip('.') + " (Tabella 12 e Figure 11-12)."


def gearing_intro(dati):
    g, gs = _serie(dati, 'gearing')
    if g.valida and gs.valida and g.vN > gs.vN * 1.5:
        return ("L'elemento che maggiormente differenzia il profilo patrimoniale è il Gearing, che "
                "segnala un ricorso al capitale di terzi significativamente superiore al comparto.")
    if g.valida and gs.valida and g.vN > gs.vN:
        return "Il Gearing segnala un ricorso al capitale di terzi superiore al comparto."
    return "Il Gearing segnala un ricorso al capitale di terzi contenuto rispetto al comparto."


def gearing_paragrafo_1(dati):
    g, gs = _serie(dati, 'gearing')
    if not g.valida:
        return ''
    frasi = [f"Il Gearing è pari {na(g.vN, '%', 'a')} nel {g.aN}."]
    passi = passo_per_passo(g, '%')
    if passi:
        frasi.append("L'indicatore, " + passi + ".")
    if g.vN < g.v0:
        tendenza = f"La tendenza complessiva della società è quindi di riduzione rispetto al {g.a0}"
    elif g.vN > g.v0:
        tendenza = f"La tendenza complessiva della società è quindi di aumento rispetto al {g.a0}"
    else:
        tendenza = f"Il valore è quindi invariato rispetto al {g.a0}"
    if gs.valida:
        if g.vN > gs.vN * 1.5:
            tendenza += ", ma il livello rimane molto superiore alla mediana settoriale"
        elif g.vN > gs.vN:
            tendenza += ", con un livello ancora superiore alla mediana settoriale"
        else:
            tendenza += ", con un livello inferiore alla mediana settoriale"
    frasi.append(tendenza + ".")
    return " ".join(frasi)


def gearing_paragrafo_2(dati):
    g, gs = _serie(dati, 'gearing')
    if not (g.valida and gs.valida):
        return ''
    frasi = []
    if g.vN > gs.vN:
        frasi.append(f"Nel {g.aN} il Gearing aziendale supera di {pp(g.vN - gs.vN)} la mediana del settore, "
                     f"pari {na(gs.vN, '%', 'a')}.")
    else:
        frasi.append(f"Nel {g.aN} il Gearing aziendale è inferiore di {pp(gs.vN - g.vN)} alla mediana del "
                     f"settore, pari {na(gs.vN, '%', 'a')}.")
    lineare = g.sale_sempre or g.scende_sempre
    andamento = ("lineare" if lineare else "non lineare") + (" ma complessivamente decrescente"
                                                               if not lineare and g.vN < g.v0 else
                                                               " ma complessivamente crescente"
                                                               if not lineare and g.vN > g.v0 else "")
    sett_mov = ("diminuisce in modo continuo" if gs.scende_sempre else
                "cresce in modo continuo" if gs.sale_sempre else
                ("diminuisce" if gs.vN < gs.v0 else "cresce"))
    frasi.append(f"Il confronto storico è particolarmente significativo: la mediana settoriale {sett_mov}, "
                 f"{na(gs.v0, '%', 'da')} nel {gs.a0} {na(gs.vN, '%', 'a')} nel {gs.aN}, mentre per la "
                 f"società il dato passa {na(g.v0, '%', 'da')} {na(g.vN, '%', 'a')}, con un andamento "
                 f"{andamento}.")
    rapp0 = g.v0 / gs.v0 if gs.v0 else None
    rappN = g.vN / gs.vN if gs.vN else None
    if g.vN < g.v0 and gs.vN < gs.v0 and rapp0 and rappN and rappN > rapp0:
        frasi.append(f"Di conseguenza, pur avendo ridotto la propria leva rispetto al {g.a0}, la società non "
                     f"ha seguito il ritmo di riduzione osservato nel comparto e il differenziale relativo si "
                     f"è ampliato.")
    if g.vN > gs.vN * 1.5:
        frasi.append("Questo rende il Gearing il principale elemento di distanza strutturale rispetto al "
                     "settore")
        return " ".join(frasi) + " (Tabella 13 e Figure 13-14)."
    return " ".join(frasi).rstrip('.') + " (Tabella 13 e Figure 13-14)."


def patr_conclusioni(dati):
    """Quattro capoversi di chiusura dell'Equilibrio Patrimoniale."""
    s1, t1 = _serie(dati, 'strut1')
    s2, t2 = _serie(dati, 'strut2')
    g, gs = _serie(dati, 'gearing')
    copertura = s1.valida and s2.valida and s1.vN > 1 and s2.vN > 1
    leva_alta = g.valida and gs.valida and g.vN > gs.vN
    anno = s1.aN if s1.valida else 2024

    p1 = f"Nel {anno} l'Equilibrio Patrimoniale appartiene alla classe “{dati['rating_patr']}”."
    if copertura and leva_alta:
        p1 += (" Gli indici di struttura superiori all'unità confermano la copertura delle immobilizzazioni "
               "mediante fonti durevoli, mentre il Gearing elevato rappresenta il principale elemento di "
               "attenzione.")
    elif copertura:
        p1 += (" Gli indici di struttura superiori all'unità confermano la copertura delle immobilizzazioni "
               "mediante fonti durevoli, con un ricorso al debito contenuto.")
    else:
        p1 += " La copertura delle immobilizzazioni mediante fonti durevoli non è piena."

    # Il quadriennio: si confrontano gli spostamenti veri, non si presume che il
    # settore abbia fatto meglio.
    p2 = ''
    if s1.valida and t1.valida and g.valida and gs.valida:
        struttura_su = s1.vN > s1.v0 and s2.vN > s2.v0
        leva_giu = g.vN < g.v0
        pezzi = []
        if struttura_su and leva_giu:
            pezzi.append(f"La lettura del quadriennio mostra un miglioramento degli indici di struttura e una "
                         f"riduzione del Gearing rispetto al {g.a0}")
        elif struttura_su:
            pezzi.append("La lettura del quadriennio mostra un miglioramento degli indici di struttura")
        elif leva_giu:
            pezzi.append(f"La lettura del quadriennio mostra una riduzione del Gearing rispetto al {g.a0}")
        recupero_struttura = (t1.vN - s1.vN) < (t1.v0 - s1.v0) and (t2.vN - s2.vN) < (t2.v0 - s2.v0)
        sett_leva_giu_piu = (gs.var or 0) < (g.var or 0)
        dettaglio = []
        if recupero_struttura and s1.vN < t1.vN:
            dettaglio.append("sugli indici di struttura la società riduce la distanza dal settore, pur "
                             "restando al di sotto delle mediane")
        if leva_giu and sett_leva_giu_piu:
            dettaglio.append("sul Gearing, invece, il settore ha registrato una riduzione della propria "
                             "mediana molto più marcata")
        if pezzi:
            frase = pezzi[0]
            if dettaglio:
                frase += "; " + "; ".join(dettaglio)
            p2 = frase + "."
            if dettaglio and sett_leva_giu_piu:
                p2 += (" La posizione patrimoniale della società appare quindi migliorata in termini "
                       "assoluti, ma meno favorevole in termini relativi per quanto riguarda l'indebitamento.")

    p3 = ''
    if copertura and leva_alta:
        p3 = ("Il profilo patrimoniale combina quindi una buona correlazione temporale tra fonti e impieghi "
              "di lungo periodo con una leva finanziaria ancora elevata. Il dato suggerisce di monitorare nel "
              "tempo sia la composizione delle fonti sia la sostenibilità del relativo costo, evitando di "
              "affidare ai soli indici di struttura l'intera lettura dell'equilibrio patrimoniale.")
    p4 = ''
    if copertura and leva_alta:
        p4 = ("Nel complesso, la struttura patrimoniale appare adeguata sotto il profilo della copertura "
              "delle immobilizzazioni, mentre il livello di indebitamento costituisce il principale elemento "
              "differenziante rispetto al settore.")
    elif copertura:
        p4 = ("Nel complesso, la struttura patrimoniale appare adeguata sia sotto il profilo della copertura "
              "delle immobilizzazioni sia per il livello di indebitamento.")
    return [p for p in (p1, p2, p3, p4) if p]


# --- Analisi degli Equilibri: area finanziaria ----------------------------

def fin_intro_1(dati):
    cr, _ = _serie(dati, 'cr')
    qr, _ = _serie(dati, 'qr')
    ro, ros = _serie(dati, 'rotazione')
    if not (cr.valida and qr.valida):
        return ''
    frasi = []
    copertura = ("una copertura delle passività correnti sempre superiore all'unità" if cr.sempre_sopra(1)
                 else "una copertura delle passività correnti non sempre superiore all'unità")
    rotazione = ""
    if ro.valida and ros.valida:
        sopra, _ = sopra_in_tutti(ro, ros)
        rotazione = (" e un'elevata capacità di rotazione del capitale investito"
                     if len(sopra) >= len(ro.punti) - 1 else "")
    frasi.append(f"{periodo_anni(cr)} l'Equilibrio Finanziario evidenzia {copertura}{rotazione}.")
    sotto_picco = cr.picco and qr.picco and cr.amax == qr.amax == cr.aP
    if sotto_picco:
        f = (f"Nel {cr.aN}, tuttavia, Current Ratio e Quick Ratio risultano inferiori ai rispettivi massimi "
             f"del {cr.aP}")
        if ro.valida and (ro.var_ultimo or 0) > MARCATO:
            f += ", mentre la rotazione del capitale investito aumenta in modo molto marcato"
        frasi.append(f + ".")
    return " ".join(frasi)


def fin_intro_2(dati):
    cr, crs = _serie(dati, 'cr')
    qr, qrs = _serie(dati, 'qr')
    ro, ros = _serie(dati, 'rotazione')
    if not (crs.valida and qrs.valida):
        return ''
    frase = (f"Il confronto con il settore mostra una dinamica differenziata: la mediana del Current Ratio "
             f"{'cresce' if crs.vN > crs.v0 else 'scende'} da {n(crs.v0)} a {n(crs.vN)} e quella del Quick "
             f"Ratio da {n(qrs.v0)} a {n(qrs.vN)}")
    if ros.valida:
        frase += f", mentre la mediana della rotazione passa da {n(ros.v0)} a {n(ros.vN)}"
    frasi = [frase + "."]
    if cr.valida and cr.vN > 1 and cr.vN < cr.vP and crs.vN > crs.v0:
        frasi.append(f"La società mantiene quindi indicatori di breve periodo complessivamente solidi, ma "
                     f"il ridimensionamento del {cr.aN} va valutato alla luce del contemporaneo miglioramento "
                     f"dei benchmark settoriali.")
    return " ".join(frasi)


def current_ratio_paragrafi(dati):
    """Tre capoversi: valore 2024, confronto con il settore, lettura finale."""
    cr, crs = _serie(dati, 'cr')
    if not cr.valida:
        return ['', '', '']
    p1 = (f"Nel {cr.aN} il Current Ratio è pari a {n(cr.vN)}. Il valore indica che le attività correnti "
          f"complessive sono pari a {n(cr.vN)} volte le passività correnti"
          + (" e mostra quindi una buona copertura dei debiti a breve termine." if cr.vN > 1.2 else
             " e mostra quindi una copertura dei debiti a breve termine." if cr.vN > 1 else
             ": la copertura dei debiti a breve termine non è piena."))
    if cr.sempre_sopra(1) and cr.picco:
        p1 += (f" L'indicatore è rimasto superiore all'unità in tutti gli anni, raggiungendo il massimo "
               f"di {n(cr.vmax)} nel {cr.amax} prima di ridursi nel {cr.aN}.")
    elif cr.sempre_sopra(1):
        p1 += " L'indicatore è rimasto superiore all'unità in tutti gli anni."
    p2 = ''
    if crs.valida:
        passi = [b - a for (_, a), (_, b) in zip(crs.punti, crs.punti[1:])]
        ultimo_piu_ampio = passi and passi[-1] == max(passi) and passi[-1] > 0
        p2 = (f"La mediana settoriale, invece, passa da {n(crs.v0)} nel {crs.a0} a {n(crs.vN)} nel {crs.aN}"
              + (", con un incremento più evidente proprio nell'ultimo esercizio." if ultimo_piu_ampio else "."))
        prossimo = abs(cr.vN - crs.vN) / abs(crs.vN) * 100 <= VICINO if crs.vN else False
        if cr.vN < cr.vP and crs.vN > crs.vP:
            p2 += (f" Di conseguenza, il valore aziendale di {n(cr.vN)} nel {cr.aN} "
                   + ("rimane prossimo al benchmark" if prossimo else "si allontana dal benchmark")
                   + ", ma la distanza relativa è meno favorevole rispetto agli esercizi precedenti: la "
                     "società ha ridotto il proprio Current Ratio mentre il settore lo ha aumentato")
        p2 = p2.rstrip('.') + " (Tabella 14 e Figura 15)."
    p3 = ''
    if cr.vN > 1 and cr.vN < cr.vP:
        p3 = (f"Nel complesso la situazione di breve periodo rimane quindi equilibrata, ma il calo del Current "
              f"Ratio nel {cr.aN} merita di essere seguito nei successivi esercizi, soprattutto in relazione "
              f"all'evoluzione dei debiti correnti e delle attività correnti (Tabella 14 e Figura 16).")
    elif cr.vN > 1:
        p3 = ("Nel complesso la situazione di breve periodo rimane quindi equilibrata (Tabella 14 e Figura 16).")
    else:
        p3 = ("Nel complesso la situazione di breve periodo richiede attenzione: le attività correnti non "
              "coprono le passività di pari scadenza (Tabella 14 e Figura 16).")
    return [p1, p2, p3]


def quick_ratio_coda(dati):
    """Continua 'Passando all'analisi del Quick Ratio (...) registra un valore pari a 1,59'."""
    qr, qrs = _serie(dati, 'qr')
    cr, _ = _serie(dati, 'cr')
    if not qr.valida:
        return '.'
    frasi = [f", contro {na(qrs.vN, '', None)} della mediana settoriale." if qrs.valida else "."]
    if qr.vN > 1:
        frasi.append("Le attività correnti al netto delle rimanenze sono sufficienti a coprire le "
                     "passività correnti.")
    else:
        frasi.append("Le attività correnti al netto delle rimanenze non bastano a coprire le passività "
                     "correnti.")
    if qr.sempre_sopra(1):
        analogo = "Analogamente all'indicatore precedente, anche il" if cr.valida and cr.sempre_sopra(1) else "Il"
        f = f"{analogo} Quick Ratio è rimasto superiore all'unità {tutto_il_periodo(qr, 'per')}"
        if qr.picco:
            f += f", raggiungendo {n(qr.vmax)} nel {qr.amax} e riducendosi a {n(qr.vN)} nel {qr.aN}"
        frasi.append(f)
    return (frasi[0] + " " + " ".join(frasi[1:])).rstrip('.') + " (Tabella 15 e Figure 17-18)."


def rotazione_paragrafo(dati):
    """Continua 'Pur tenendo conto delle fisiologiche differenze ...'."""
    ro, ros = _serie(dati, 'rotazione')
    if not (ro.valida and ros.valida):
        return "la performance della società va letta insieme alla dinamica del settore (Tabella 16 e Figura 19)."
    frasi = []
    if ro.vN > ros.vN:
        frasi.append("la performance della società si distingue per valori elevati dell'Indice di Rotazione.")
    else:
        frasi.append("la performance della società si colloca al di sotto della mediana dell'Indice di Rotazione.")
    f = f"L'indicatore {confronto_anni_benchmark(ro, ros)}"
    if (ro.var_ultimo or 0) > MARCATO:
        f += (f", ma il valore {ro.aN} rappresenta un cambiamento molto marcato rispetto "
              f"{na(ro.vP, '', 'a')} del {ro.aP}")
    frasi.append(f + ".")
    if (ros.var_ultimo or 0) > 0 and (ro.var_ultimo or 0) > MARCATO:
        frasi.append(f"Anche il settore registra un incremento nell'ultimo esercizio, passando da una mediana "
                     f"di {n(ros.vP)} nel {ros.aP} a {n(ros.vN)} nel {ros.aN}, ma l'incremento aziendale è "
                     f"molto più accentuato.")
    if (ro.var_ultimo or 0) > MARCATO:
        frasi.append("Un aumento così rilevante non è automaticamente sinonimo di maggiore efficienza e "
                     "pertanto merita un approfondimento, al fine di verificare se derivi da un incremento del "
                     "Valore della Produzione, da una riduzione del capitale investito o da entrambe le dinamiche")
        return " ".join(frasi) + " (Tabella 16 e Figura 19)."
    return " ".join(frasi).rstrip('.') + " (Tabella 16 e Figura 19)."


def fin_conclusioni(dati):
    cr, crs = _serie(dati, 'cr')
    qr, qrs = _serie(dati, 'qr')
    ro, ros = _serie(dati, 'rotazione')
    positiva = cr.valida and qr.valida and cr.vN > 1 and qr.vN > 1
    p1 = (f"Il giudizio complessivo, per cui l'impresa appartiene alla classe “{dati['rating_fin']}”, "
          f"riflette quindi una configurazione finanziaria di breve periodo "
          f"{'positiva' if positiva else 'da presidiare'} secondo le metriche considerate.")
    calo_liq = cr.valida and qr.valida and cr.vN < cr.vP and qr.vN < qr.vP
    bench_su = crs.valida and qrs.valida and crs.vN > crs.vP and qrs.vN > qrs.vP
    salto_rot = ro.valida and (ro.var_ultimo or 0) > MARCATO
    if calo_liq or salto_rot:
        f = " La lettura temporale richiede tuttavia di distinguere i diversi indicatori: "
        pezzi = []
        if calo_liq:
            pezzi.append(f"Current Ratio e Quick Ratio si riducono rispetto al {cr.aP}"
                         + (", mentre i rispettivi benchmark settoriali aumentano" if bench_su else ""))
        if salto_rot:
            pezzi.append(f"la rotazione del capitale investito registra nel {ro.aN} un incremento eccezionale"
                         + (", molto superiore a quello osservato nel settore" if ros.valida and
                            (ros.var_ultimo or 0) < (ro.var_ultimo or 0) / 2 else ""))
        f += "; al contrario, ".join(pezzi) if len(pezzi) == 2 else pezzi[0]
        p1 += f + "."
    p2 = (f"Nel complesso, la condizione finanziaria di breve termine rimane "
          f"{'favorevole' if positiva else 'da monitorare'} secondo gli indicatori considerati.")
    if calo_liq:
        p2 += (f" Al tempo stesso, la riduzione di Current Ratio e Quick Ratio rispetto al {cr.aP} suggerisce di "
               f"monitorare l'evoluzione del margine di copertura"
               + (f", soprattutto perché il settore ha evidenziato nel {cr.aN} un rafforzamento delle proprie "
                  f"mediane." if bench_su else "."))
    if salto_rot:
        p2 += (" La forte crescita della rotazione del capitale investito merita invece un'analisi separata "
               "delle componenti del capitale investito, poiché il differenziale rispetto al settore è molto "
               "ampio e non consente, da solo, di concludere che si tratti di un miglioramento strutturale "
               "dell'efficienza.")
    return [p1, p2]


def testi_report(dati):
    """Tutti i testi nuovi del report, pronti per il context del template."""
    t = {
        'rev_rating_tot': descr_rating_tot(dati),
        'rev_intro_margini_coda': intro_margini_coda(dati),
        'rev_rating_combinato': sintesi_rating_combinato(dati),
        'rev_sintesi_eco': sintesi_sul_piano_economico(dati),
        'rev_sintesi_patr': sintesi_sul_piano_patrimoniale(dati),
        'rev_sintesi_fin': sintesi_sul_piano_finanziario(dati),
        'rev_sintesi_finale': sintesi_finale(dati),
        'rev_ebitda_diff': ebitda_differenza(dati),
        'rev_ebitda_comp': ebitda_composizione(dati),
        'rev_ebitda_trend': ebitda_trend(dati),
        'rev_ebit': ebit_paragrafo(dati),
        'rev_profitto': profitto_paragrafo(dati),
        'rev_profitto_struttura': profitto_struttura(dati),
        'rev_eco_concl_1': eco_conclusione_1(dati),
        'rev_eco_concl_2': eco_conclusione_2(dati),
        'rev_patr_intro_1': patr_intro_1(dati),
        'rev_patr_intro_2': patr_intro_2(dati),
        'rev_str1': struttura1_paragrafo(dati),
        'rev_str2': struttura2_paragrafo(dati),
        'rev_gearing_intro': gearing_intro(dati),
        'rev_gearing_1': gearing_paragrafo_1(dati),
        'rev_gearing_2': gearing_paragrafo_2(dati),
        'rev_fin_intro_1': fin_intro_1(dati),
        'rev_fin_intro_2': fin_intro_2(dati),
        'rev_qr_coda': quick_ratio_coda(dati),
        'rev_rotazione': rotazione_paragrafo(dati),
    }
    for i, testo in enumerate(patr_conclusioni(dati) + ['', '', '', ''], start=1):
        if i > 4:
            break
        t[f'rev_patr_concl_{i}'] = testo
    for i, testo in enumerate(current_ratio_paragrafi(dati), start=1):
        t[f'rev_cr_{i}'] = testo
    for i, testo in enumerate(fin_conclusioni(dati), start=1):
        t[f'rev_fin_concl_{i}'] = testo
    t['rev_bullet'] = {k: bullet_indicatore(k, dati['az'].get(k, {}), dati['sett'].get(k, {}))
                       for k in NOMI_ART}
    return t
