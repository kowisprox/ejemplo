import json, math
A=[50,1,100,20,19,40,99]; n=len(A); k=max(A)
ops={}
# COUNTING
C=[0]*(k+1)
for v in A: C[v]+=1
occ={i:C[i] for i in range(k+1) if C[i]}
print("COUNTING  k=",k,"len(C)=",len(C)); print(" conteo (ocupados):",occ)
for i in range(1,k+1): C[i]+=C[i-1]
acc={i:C[i] for i in sorted(occ)}
print(" acumulado en ocupados:",acc)
B=[None]*n; place=[]
for j in range(n-1,-1,-1):
    C[A[j]]-=1; B[C[A[j]]]=A[j]; place.append((j,A[j],C[A[j]]))
print(" colocacion (j,x,pos):",place); print(" B=",B)
ops_count=(k+1)+n+k+n+n  # init, count, accumulate, place, (copy/read)
print(" ops: init",k+1,"+ contar",n,"+ acumular",k,"+ colocar",n,"=",(k+1)+n+k+n)
# RADIX
R=A[:]; passes=[]
d=len(str(max(A)))
for p in range(d):
    bk=[[] for _ in range(10)]
    for x in R: bk[(x//10**p)%10].append(x)
    R=[x for b in bk for x in b]
    passes.append({"p":p+1,"buckets":{i:b for i,b in enumerate(bk) if b},"out":R[:]})
    print("RADIX pasada",p+1,{i:b for i,b in enumerate(bk) if b},"->",R)
# BUCKET
mn,mx=min(A),max(A); m=n
bk=[[] for _ in range(m)]; idx=[]
for x in A:
    j=m*(x-mn)//(mx-mn+1); bk[j].append(x); idx.append((x,j))
print("BUCKET min,max",mn,mx,"idx",idx); print(" cubetas",bk)
srt=[sorted(b) for b in bk]; print(" ordenadas",srt," ->",[x for b in srt for x in b])
rng=[(math.ceil(j*(mx-mn+1)/m)+mn, math.ceil((j+1)*(mx-mn+1)/m)+mn-1) for j in range(m)]
print(" rangos",rng)
# mergesort comparisons
cmp=0
def ms(a):
    global cmp
    if len(a)<=1: return a
    l=ms(a[:len(a)//2]); r=ms(a[len(a)//2:]); o=[];i=j=0
    while i<len(l) and j<len(r):
        cmp+=1
        if l[i]<=r[j]: o.append(l[i]);i+=1
        else: o.append(r[j]);j+=1
    return o+l[i:]+r[j:]
ms(A); print("mergesort comparaciones reales:",cmp,"n log2 n =",round(n*math.log2(n),1))
print("counting total n+k ... =",n+k, " con init k+1:", (k+1)+n+k+n)
json.dump({"A":A,"occ":occ,"acc":acc,"place":place,"passes":passes,"idx":idx,"bk":bk,"rng":rng,"cmp":cmp},open("trace.json","w"))
