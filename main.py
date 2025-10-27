'''''
k=0
alf='acgt'
for a1 in alf:
    for a2 in alf:
        for a3 in alf:
            for a4 in alf:
                for a5 in alf:
                    s=a1+a2+a3+a4+a5
                    if s.count('a')==2:
                        k+=1
print(k)
'''

from intertools import *
p=product ('acgt', repeat=5)
k=0
for x in p:
        if x.count ('a') ==2: k+=1
print (k)