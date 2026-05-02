import pandas as pd

#variables
a=10
name="priyanshu"

#printing variables
print(a)
print(name)

#lists
list1=[1,2,4,23,5,7]
print(list1)

list2=["pewfrwewe",12,1234,"axis"]
print(list2)

#dictionaries
dict1={"name": ["priyanshu","alok","sunita","parul"],
       "age": [20,21,22,23],
       "city": ["bhopal","delhi","mumbai","koldas"]}
print(dict1)

#dataframes
data={"client":["a","b","c","d","e"],
      "engagement":[10,20,60,20,90],
      "revenue": [1000,2000,3000,6000,90000]}
df=pd.DataFrame(data)

