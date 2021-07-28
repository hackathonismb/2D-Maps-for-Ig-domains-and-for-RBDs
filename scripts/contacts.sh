sed -e 's/},/\n/g' 6m0j_contact_map.json >file1.txt
cnt=1
for i in `ls | grep file1.txt | sort -V`
do
	j="out_1"$cnt".txt"
	sed -n '3,$'p $i > $j
	cnt=$((cnt+1))
done
cat $j > f1.txt
cat f1.txt |grep "0F0" | sort -V > test1.txt
cat f1.txt |grep "FF0000" | sort -V > test2.txt
cat f1.txt |grep "C8C8C8" | sort -V > test3.txt
cat f1.txt |grep "0FF" | sort -V > test4.txt
cat f1.txt |grep "F00" | sort -V > test5.txt
cat f1.txt |grep "00F" | sort -V > test6.txt
cat f1.txt |grep "888" | sort -V > test7.txt
cat f1.txt |grep "FFA500" | sort -V > test8.txt
cat f1.txt |grep "006400" | sort -V > test9.txt
cat f1.txt |grep "FFF" | sort -V > test10.txt
cat f1.txt |grep "FOF" | sort -V > test11.txt
cat f1.txt |grep "OOF" | sort -V > test12.txt
