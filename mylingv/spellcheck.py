#!/usr/bin/python2
# -*- coding: utf-8
import re, struct, os, time
from functools import partial

commonDict = set()
dawg = []
with open(os.path.join(os.path.dirname(__file__), "numbers.dawg"), "rb") as f:
    for item in iter(partial(f.read, 8), ''):
       dawg.append(struct.unpack_from("<BBHi", item))
rootlen = len(dawg)
for item in dawg:
    if item[3] > 0:
        rootlen = item[3]
        break

class SearchState: pass
class DawgSearch:
    def __init__(self):
        self.states = []
        self.state = SearchState()
        self.state.node=0
        self.state.len=rootlen
        self.state.isTerminal=False
        self.currentPrefix = ""
    def _bisect(self, x):
        lo = self.state.node
        hi = lo + self.state.len
        while lo < hi:
            mid = (lo+hi)//2
            val = dawg[mid][2]
            if val == x: return mid
            elif val < x: lo = mid+1
            else: hi = mid
        return None

    def update(self, char, idx = None):
        if idx == None: idx = self._bisect(ord(char))
        if idx == None:
            raise Exception('Cannot continue prefix "%s" with character "%s"' % (self.currentPrefix, char))
        self.currentPrefix += char
        self.states.append(self.state)
        self.state = SearchState()
        self.state.node = dawg[idx][3]
        self.state.len = dawg[idx][1]
        self.state.isTerminal = dawg[idx][0]>0

    def backspace(self):
        self.state = self.states.pop()
        self.currentPrefix = self.currentPrefix[:-1]

    def isFullWord(self): return self.state.isTerminal
    def canUpdate(self, char): return self._bisect(ord(char)) != None
    def nextLetters(self): return ((unichr(dawg[i][2]), i) for i in xrange(self.state.node, self.state.node+self.state.len))
    def foreachLetter(self, fn):
        for letter, idx in self.nextLetters():
            self.update(letter, idx)
            fn(letter)
            self.backspace()
    def getWords(self):
        if self.isFullWord(): yield self.currentPrefix
        for letter, idx in self.nextLetters():
            self.update(letter, idx)
            for i in self.getWords(): yield i
            self.backspace()

def search( word, maxCost ):

    currentRow = range( len(word) + 1 )

    results = []
    ds = DawgSearch()
    ds.foreachLetter(lambda letter: searchRecursive(ds, letter, word, currentRow, results, maxCost))
    return results

# This recursive helper is used by the search function above. It assumes that
# the previousRow has been filled in already.
def searchRecursive( ds, letter, word, previousRow, results, maxCost ):
    columns = len( word ) + 1
    currentRow = [ previousRow[0] + 1 ]

    # Build one row for the letter, with a column for each letter in the target
    # word, plus one for the empty string at column 0
    for column in xrange( 1, columns ):

        insertCost = currentRow[column - 1] + 1
        deleteCost = previousRow[column] + 1

        if word[column - 1] != letter:
            replaceCost = previousRow[ column - 1 ] + 1
        else:
            replaceCost = previousRow[ column - 1 ]

        currentRow.append( min( insertCost, deleteCost, replaceCost ) )
    # if the last entry in the row indicates the optimal cost is less than the
    # maximum cost, and there is a word in this trie node, then add it.
    if currentRow[-1] <= maxCost and ds.isFullWord():
        results.append((ds.currentPrefix, currentRow[-1]))

    if min( currentRow ) > maxCost: return
    # if any entries in the row are less than the maximum cost, then
    # recursively search each branch of the trie
    ds.foreachLetter(lambda letter: searchRecursive(ds, letter, word, currentRow, results, maxCost))

def fixword(word):
    if re.match("[0-9][0-9]?", word):
        return word
    if len(word)>15: return None
    if len(word)==1: return None
    word = word.lower()
    results = search(word, 1 if len(word)<5 else 2)
    if len(results) == 0: return None
    results.sort(lambda a,b:cmp(a[1],b[1]))
    if len(results)>1 and results[0][1] == results[1][1]:
        return None # Есть разные варианты коррекции для слова
    return results[0][0]

def filterText(text):
    result = u""
    for word in re.finditer(ur"[а-я0-9]+", text, re.IGNORECASE):
        word = fixword(word.group(0))
        if word != None: result += word + " ";
    return result

if __name__ == "__main__":
    import sys, time, cProfile, pstats, StringIO
    pr = cProfile.Profile()
    pr.enable()
    with open(sys.argv[1], "r") as f:
        start = time.time()
        for line in f:
            line = line.decode("utf-8")
            for word in re.finditer(ur"[а-я0-9]+", line, re.IGNORECASE):
                word = word.group(0)
                fw = fixword(word)
                if fw == None: continue
                if fw == word: sys.stdout.write(word.encode("utf-8"))
                else: sys.stdout.write(("%s -> %s" % (word, fw)).encode("utf-8"))
                sys.stdout.write(" ")
                sys.stdout.flush()
        end = time.time()
    print
    print("Spellchecked file %s, time: %f s" % (sys.argv[1], end - start))
    s = StringIO.StringIO()
    ps = pstats.Stats(pr, stream=s).sort_stats('cumulative')
    ps.print_stats()
    print s.getvalue()
