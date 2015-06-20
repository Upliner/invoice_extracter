#!/usr/bin/python2
# -*- coding: utf-8
import re, struct, os

if __name__ == "__main__":
    import sys, time, cProfile, pstats, StringIO
    pr = cProfile.Profile()
    pr.enable()
    start = time.time()

with open(os.path.join(os.path.dirname(__file__), "numbers.dawg"), "rb") as f:
    dawg = f.read()

nodes = {}
rootlen = len(dawg)
i = 0
def readNode(idx):
    global rootlen
    cachedNode = nodes.get(idx)
    if cachedNode != None: return cachedNode
    item = struct.unpack_from("<BBHi", dawg, idx*8)
    if item[3]>i and item[3]<rootlen: rootlen = item[3]
    result = (item[0]>0, unichr(item[2]), [readNode(child) for child in xrange(item[3], item[3]+item[1])])
    nodes[idx] = result
    return result
dawgRoot = []
while i < rootlen:
    dawgRoot.append(readNode(i))
    i += 1
del dawg
del nodes

def search( word, maxCost ):
    global _maxCost
    _maxCost = maxCost
    currentRow = range( len(word) + 1 )
    results = []
    for node in dawgRoot:
        searchRecursive(node, "", node[1], word, currentRow, results)
    return results

# This recursive helper is used by the search function above. It assumes that
# the previousRow has been filled in already.
def searchRecursive(node, prefix, letter, word, previousRow, results):
    global _maxCost
    columns = len( word ) + 1
    prefix+=letter
    i = len(prefix)
    currentRow = range(i, i+columns)
    lo = max(1, i-_maxCost-1)
    hi = min(i+_maxCost+1,columns)
    # Build one row for the letter, with a column for each letter in the target
    # word, plus one for the empty string at column 0
    for column in xrange(lo, hi):
        if word[column - 1] == letter:
            currentRow[column] = previousRow[ column - 1 ]
        else:
            currentRow[column] = min(
                currentRow[column - 1] + 1,
                previousRow[column] + 1,
                previousRow[column - 1] + 1)

    # if the last entry in the row indicates the optimal cost is less than the
    # maximum cost, and there is a word in this trie node, then add it.
    if currentRow[-1] <= _maxCost and node[0]:
        results.append((prefix, currentRow[-1]))
        _maxCost = currentRow[-1]
        if _maxCost == 0: return
    elif min(currentRow[lo:hi] ) > _maxCost: return
    # if any entries in the row are less than the maximum cost, then
    # recursively search each branch of the trie
    for subNode in node[2]:
        searchRecursive(subNode, prefix, subNode[1], word, currentRow, results)

def fixword(word):
    if re.match("[0-9][0-9]?$", word):
        return word
    if len(word)>15: return None
    if len(word)==1: return None
    word = word.lower()
    results = search(word, 1 if len(word)<5 else 2)
    if len(results) == 0: return None
    results.sort(lambda a,b:cmp(a[1],b[1]))
    if len(results)>1 and results[0][1] == results[1][1]:
        # Есть разные варианты коррекции для слова
        # Если это плохо распознанное число, не выкидываем его из потока
        # для избежения неправильного распознавания суммы прописью
        return word
    return results[0][0]

def filterText(text):
    result = u""
    for word in re.finditer(ur"[а-яА-Я0-9]+", text, re.IGNORECASE):
        word = fixword(word.group(0))
        if word == None:
            if len(result)>0: yield result; result = u""
        else: result += word + " ";
    if len(result)>0: yield result

if __name__ == "__main__":
    print("Loaded, time: %f s" % (time.time() - start))
    with open(sys.argv[1], "r") as f:
        for line in f:
            line = line.decode("utf-8")
            for word in re.finditer(ur"[а-яА-Я0-9]+", line, re.IGNORECASE):
                word = word.group(0)
                fw = fixword(word)
                if fw == None: continue
                if fw == word: sys.stdout.write(word.encode("utf-8"))
                else: sys.stdout.write(("%s -> %s" % (word, fw)).encode("utf-8"))
                sys.stdout.write(", ")
                sys.stdout.flush()
    end = time.time()
    print
    print("Spellchecked file %s, time: %f s" % (sys.argv[1], end - start))
    s = StringIO.StringIO()
    ps = pstats.Stats(pr, stream=s).sort_stats('cumulative')
    ps.print_stats()
    print s.getvalue()
