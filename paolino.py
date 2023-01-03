# -*- coding: utf-8 -*-

import re
import xlrd
from collections import defaultdict

INTER = "inter"

def spacify(value):
    return re.sub(r"[,./\ ]+", ' ', value).lower()

def get_stadio(value):
    parts = spacify(value).split()
    stadio = parts.pop(0)
    if stadio == 'reggio':
        for check, add in (('mirabello', None), ('giglio', 'mapei'), ('mapei', None)):
            if check in parts:
                stadio += add or check
                break
        else:
            assert False
    else:
        nome = parts.pop()
        if not (stadio in ('vigevano', 'udine', 'bergamo') or stadio in ('torino') and nome in ('olimpico', 'comunale', 'torino')):
            stadio += nome
    return stadio

MAP_SQUADRE = {
        u'casino salzburg': u'casino salisburgo',
        u'lyon': u'lione',
        u'spartak moskva': u'spartak mosca',
        u'sporting lisboa': u'sporting lisbona',
        u'olympique lyonnese': u'lione',
        u'barcelona': u'barcellona',
        u'fcb basel': u'basilea',
        u'saragoza': u'saragozza',
        u'strasbourg': u'strasburgo',
        u'hellas verona': u'verona',
        u'slavia praha': u'slavia praga',
        u'celtic': u'celtic glasgow',
        }
def map_squadra(value):
    return MAP_SQUADRE.get(value, value)

def get_squadre(value, n=1):
    if n > 1:
        parts = spacify(value).split('-', 1)
        return [map_squadra(part.strip()) for part in parts]
    else:
        return map_squadra(spacify(value))

def check(args):
    wb = xlrd.open_workbook(args.filename)
    sh_partite = wb.sheet_by_name("partite allo stadio")
    sh_stadi = wb.sheet_by_name("stadi")
    sh_squadre = wb.sheet_by_name("squadre")
    stadi = dict((get_stadio(sh_stadi.cell(n, 1).value), n) for n in range(sh_stadi.nrows) if sh_stadi.cell(n, 0).ctype == xlrd.XL_CELL_NUMBER)
    squadre = dict((get_squadre(sh_squadre.cell(n, 1).value), n) for n in range(sh_squadre.nrows) if sh_squadre.cell(n, 0).ctype == xlrd.XL_CELL_NUMBER)
    partite_stadi = defaultdict(int)
    partite_others_stadi = defaultdict(int)
    vittorie_inter = defaultdict(int)
    pareggi_inter = defaultdict(int)
    sconfitte_inter = defaultdict(int)
    last_n_stadio = last_n_stadio_used = None
    for n in range(sh_partite.nrows):
        if sh_partite.cell(n, 0).ctype == xlrd.XL_CELL_NUMBER:
            stadio = get_stadio(sh_partite.cell(n, 2).value)
            last_n_stadio = n
            assert stadio in stadi, (stadio, sh_partite.cell(n, 2).value, stadi)
        else:
            goals = sh_partite.cell(n, 3).value.strip()
            if goals:
                squadra1, squadra2 = get_squadre(sh_partite.cell(n, 2).value, 2)
                goal1, goal2 = [int(value) for value in sh_partite.cell(n, 3).value.replace('-', ' ').split()[:2]]
                ok_squadra = True
                for squadra in (squadra1, squadra2):
                    if not squadra == INTER and not squadra in squadre:
                        print(squadra.title(), ": squadra non trovata", sorted(squadre))
                        ok_squadra = False
                if ok_squadra:
                    others = False
                    if squadra1 == INTER:
                        squadra = squadra2
                        goal_inter, goal_other = goal1, goal2
                    elif squadra2 == INTER:
                        squadra = squadra1
                        goal_inter, goal_other = goal2, goal1
                    else:
                        others = True
                    counter_stadi = partite_others_stadi if others else partite_stadi
                    if others or last_n_stadio_used != last_n_stadio:
                        counter_stadi[stadi[stadio]] += 1
                        if not others:
                            last_n_stadio_used = last_n_stadio
                    if not others:
                        n_squadra = squadre[squadra]
                        if goal_inter > goal_other:
                            vittorie_inter[n_squadra] += 1
                        elif goal_inter < goal_other:
                            sconfitte_inter[n_squadra] += 1
                        else:
                            pareggi_inter[n_squadra] += 1
    for n, counter_stadi in enumerate([partite_stadi, partite_others_stadi]):
        if n == 0:
            label = 'Inter'
        else:
            label = 'altre'
        for n_stadio, n_partite in counter_stadi.items():
            n_partite_file = int(sh_stadi.cell(n_stadio, 2 + n).value or '0')
            stadio = sh_stadi.cell(n_stadio, 1).value
            if n_partite_file != n_partite:
                print("Partite %s diverse per stadio \"%s\": file=%s, calcolo=%s" % (label, stadio, n_partite_file, n_partite))
    for squadra, n_squadra in sorted(squadre.items()):
        vittorie_inter[n_squadra] += 0
        pareggi_inter[n_squadra] += 0
        sconfitte_inter[n_squadra] += 0
        vittorie, pareggi, sconfitte = vittorie_inter.pop(n_squadra), pareggi_inter.pop(n_squadra), sconfitte_inter.pop(n_squadra)
        vittorie_file, pareggi_file, sconfitte_file = [int(sh_squadre.cell(n_squadra, n).value or '0') for n in (7, 8, 9)]
        squadra = sh_squadre.cell(n_squadra, 1).value
        if vittorie_file != vittorie:
            print("Vittorie diverse per squadra \"%s\": file=%s, calcolo=%s" % (squadra, vittorie_file, vittorie))
        if pareggi_file != pareggi:
            print("Pareggi diverse per squadra \"%s\": file=%s, calcolo=%s" % (squadra, pareggi_file, pareggi))
        if sconfitte_file != sconfitte:
            print("Sconfitte diverse per squadra \"%s\": file=%s, calcolo=%s" % (squadra, sconfitte_file, sconfitte))
        totale = int(sh_squadre.cell(n_squadra, 6).value or '0')
        if sconfitte_file + pareggi_file + vittorie_file != totale:
            print("Totali partite diverse per squadra \"%s\": file=%s, calcolo=%s" % (squadra, totale, sconfitte_file + pareggi_file + vittorie_file))
    print("Check completo")
    assert not vittorie_inter
    assert not pareggi_inter
    assert not sconfitte_inter

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description='Check file Paolino',
            formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument('-f', '--filename',
          help='nome del file excel')
    args = parser.parse_args()
    check(args)
