#!/usr/bin/python2
import re, struct, sys

with open(sys.argv[1], "rb") as f:
    dawg = f.read()

nodes = {}
rootlen = len(dawg)//8
i = 0
def readNode(idx):
    global rootlen
    cachedNode = nodes.get(idx)
    if cachedNode != None: cachedNode[1] += 1; return
    item = struct.unpack_from("<BBHi", dawg, idx*8)
    if item[3]>i and item[3]<rootlen: rootlen = item[3]
    childs = range(item[3], item[3]+item[1])
    for c in childs: readNode(c)
    nodes[idx] = [(item[0]>0, unichr(item[2]), childs), 1]

dawgRoot = []
while i < rootlen:
    readNode(i)
    i += 1

def output(s):
   sys.stdout.write(s.encode("utf-8") + "\n")
tails = {}
def nodeStr(idx):
    tail = tails.get(idx)
    if tail != None: return "_tail%i" % tail
    item = nodes[idx][0]
    result = u"(%s,u'%s',%s)" % (item[0], item[1], childsStr(item[2]))
    if nodes[idx][1] > 1:
        tails[idx] = len(tails)+1
        t = "_tail%i" % len(tails)
        output(t + "=" + result)
        return t
    return result

def childsStr(arr):
    if len(arr) == 0: return "()"
    return "[" + ",".join(nodeStr(idx) for idx in arr) + "]"

output("dawg = " + childsStr(range(rootlen)))
