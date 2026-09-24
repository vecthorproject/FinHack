"""
Criteri robusti per individuare i valori anomali delle nove variabili.

Due strade, entrambe scelte perche' non presuppongono distribuzioni simmetriche:
i percentili estremi, e il boxplot adattato all'asimmetria di Hubert e
Vandervieren (An adjusted boxplot for skewed distributions, CSDA 2008), che
allarga o restringe i baffi in base al medcouple, l'indice di asimmetria
robusto di Brys, Hubert e Struyf (2004).

Il boxplot classico di Tukey e la regola della deviazione assoluta mediana
(mediana +/- k*MAD) danno per scontata una distribuzione quasi simmetrica: su
indici di bilancio, che hanno code lunghissime da un lato solo, dichiarano
anomala meta' del campione. Il medcouple misura l'asimmetria e la corregge.
"""

import numpy as np

# Coefficiente dei baffi (1,5 nel boxplot classico) e gli esponenti della
# correzione per asimmetria proposti nel lavoro del 2008.
COEFFICIENTE_BAFFI = 1.5
ESPONENTE_SINISTRA = -4.0
ESPONENTE_DESTRA = 3.0

# Oltre questa numerosita' il medcouple si stima su un sottoinsieme: la formula
# confronta tutte le coppie, e la memoria cresce col quadrato dei valori.
MASSIMO_PER_MEDCOUPLE = 4000


def _valori_puliti(valori):
    serie = np.asarray(valori, dtype=float).ravel()
    return serie[np.isfinite(serie)]


def medcouple(valori, massimo=MASSIMO_PER_MEDCOUPLE, seme=0):
    """
    Asimmetria robusta, fra -1 e +1: positiva se la coda lunga sta a destra.

    E' la mediana del nucleo h calcolato su tutte le coppie di osservazioni che
    stanno una sotto e una sopra la mediana.
    """
    serie = _valori_puliti(valori)
    if serie.size < 3:
        return 0.0
    if serie.size > massimo:
        serie = np.random.default_rng(seme).choice(serie, size=massimo, replace=False)

    centrata = np.sort(serie - np.median(serie))
    sotto = centrata[centrata <= 0.0]          # crescente: gli zeri in fondo
    sopra = centrata[centrata >= 0.0]          # crescente: gli zeri in testa
    if sotto.size == 0 or sopra.size == 0:
        return 0.0

    denominatore = sopra[:, None] - sotto
    with np.errstate(invalid='ignore', divide='ignore'):
        # dove il denominatore e' zero (valori pari alla mediana) il nucleo
        # arriva come 0/0 e viene sostituito subito dopo
        nucleo = (sopra[:, None] + sotto) / denominatore
    # I valori esattamente pari alla mediana darebbero 0/0: al loro posto va il
    # nucleo previsto dalla definizione per i valori ripetuti (-1, 0, +1 secondo
    # la posizione rispetto all'antidiagonale).
    pari_alla_mediana = int(np.sum(sotto == 0.0))
    if pari_alla_mediana:
        blocco = np.ones((pari_alla_mediana, pari_alla_mediana)) - np.eye(pari_alla_mediana)
        blocco -= 2 * np.triu(blocco)
        nucleo[:pari_alla_mediana, -pari_alla_mediana:] = np.fliplr(blocco)
    nucleo = nucleo[np.isfinite(nucleo)]
    return float(np.median(nucleo)) if nucleo.size else 0.0


def soglie_percentili(valori, percentile):
    """Le due soglie del criterio dei percentili estremi (1° e 99°, per esempio)."""
    serie = _valori_puliti(valori)
    if serie.size == 0:
        return None, None
    return (float(np.quantile(serie, percentile / 100)),
            float(np.quantile(serie, 1 - percentile / 100)))


def soglie_boxplot_adattato(valori, coefficiente=COEFFICIENTE_BAFFI):
    """
    Le due soglie del boxplot adattato all'asimmetria.

    Con mc >= 0:  [Q1 - c*e^(-4*mc)*IQR,  Q3 + c*e^(3*mc)*IQR]
    Con mc < 0:   [Q1 - c*e^(-3*mc)*IQR,  Q3 + c*e^(4*mc)*IQR]
    Se i dati sono simmetrici (mc = 0) si torna al boxplot di Tukey.
    """
    serie = _valori_puliti(valori)
    if serie.size < 4:
        return None, None
    q1, q3 = np.quantile(serie, 0.25), np.quantile(serie, 0.75)
    scarto = q3 - q1
    if scarto <= 0:
        return None, None
    mc = medcouple(serie)
    if mc >= 0:
        sinistra, destra = ESPONENTE_SINISTRA * mc, ESPONENTE_DESTRA * mc
    else:
        sinistra, destra = ESPONENTE_DESTRA * mc, ESPONENTE_SINISTRA * mc
    basso = q1 - coefficiente * np.exp(sinistra) * scarto
    alto = q3 + coefficiente * np.exp(destra) * scarto
    return float(basso), float(alto)


def soglie(valori, criterio, percentile=1.0, coefficiente=COEFFICIENTE_BAFFI):
    """Le soglie del criterio scelto: 'percentili' oppure 'boxplot'."""
    if criterio == 'boxplot':
        return soglie_boxplot_adattato(valori, coefficiente)
    return soglie_percentili(valori, percentile)
