Những hàm excel phổ biến nhất trong kế toán hiện nay
====================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/nhung-ham-excel-pho-bien-nhat-trong-ke-toan-hien-nay/)Excel được coi là một phần không thể thiếu trong kế toán doanh nghiệp. Việc sử dụng excel thành thạo và nắm rõ các hàm trong excel gần như là một yêu cầu bắt buộc với các nhân viên kế toán. Hơn nữa, excel lại có rất nhiều hàm nên việc ghi nhớ được tất …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Excel được coi là một phần không thể thiếu trong kế toán doanh nghiệp. Việc sử dụng excel thành thạo và nắm rõ các hàm trong excel gần như là một yêu cầu bắt buộc với các nhân viên kế toán. Hơn nữa, excel lại có rất nhiều hàm nên việc ghi nhớ được tất cả là một điều khó thực hiên. Để giúp các bạn trong công việc hàng ngày, bài viết dưới đây sẽ cung cấp những hàm excel phổ biến và thông dụng nhất hiện nay.


![phan-mem-excel](https://hddtvn.com/wp-content/uploads/2021/01/phan-mem-excel.jpg)


1. Hàm SUM, AVERAGE và COUNT
----------------------------


Cả 3 hàm này đều được thực hiện trên một dãy dữ liệu được chọn.


### **Hàm SUM**




* Tính tổng tất cả các số trong một vùng dữ liệu được chọn.

* Câu lệnh: =SUM(Number1, Number2…). Trong đó Number1, Number2…là các số cần tính tổng.

* Ví dụ: =SUM(4,7,8) => Kết quả =19.



### **Hàm AVERAGE**




* Tính trung bình cộng của các số trong một vùng dữ liệu được chọn.

* Câu lệnh: =AVERAGE(Number1, Number2…). Trong đó Number1, Number2…là các số cần tính trung bình cộng.

* Ví dụ: =AVERAGE(3,5,7) => Kết quả =5.



**[>>> Cách sử dụng hàm Average để tính trung bình cộng trong Excel](#)**


### **Hàm COUNT**




* Hàm dùng với mục đích đếm số lượng các ô có chứa dữ liệu kiểu số trong dãy.

* Câu lệnh: =COUNT(Value1, Value2, Value3,…)

* Ví dụ: =COUNT(A6:A9) nghĩa là đếm số ô chứa số trong các ô từ A6 đến A9.



>>>**[Hướng dẫn sử dụng các hàm COUNT, COUNTA, COUNTBLACK để đếm số ô trong Excel](#)**


2. Hàm VLOOKUP và HLOOKUP
-------------------------


Cả 2 hàm đều là hàm tìm kiếm nhưng cách tìm kiếm của 2 hàm lại khác nhau.


### **Hàm VLOOKUP**




* Hàm tìm kiếm theo cột.

* Câu lệnh: =VLOOKUP(Lookup Value, Table Array, Col idx num, [range lookup]).



![vi-du-ham-vlookup-trong-excel](https://hddtvn.com/wp-content/uploads/2021/01/vi-du-ham-vlookup-trong-excel.png)Ví dụ hàm VLOOKUP trong excel
[**>>>Hướng dẫn dùng hàm VLOOKUP trong Excel**](#)


### **Hàm HLOOKUP**




* Hàm tìm kiếm theo hàng.

* Câu lệnh: =HLOOKUP(Lookup Value, Table Array, Col idx num, [range lookup]).



*Trong đó, các tham số:*




* **Lookup Value:** Giá trị dùng để dò tìm.

* **Table Array:** Bảng dữ liệu dùng để so sánh. Vùng dữ liệu này phải là tham chiếu tuyệt đối.

* **Cox idx num:** thứ tự cột/hàng dữ liệu mà bạn muốn lấy trong phép so sánh.

* **Range lookup:** phạm vi tìm kiếm với giá trị tìm kiếm là chính xác hoặc giá trị gần đúng. X=0 là dò tìm một cách chính xác, X=1 là dò tìm một cách tương đối.



[**>>>Cách sử dụng hàm HLOOKUP trong Excel**](#)


>> [Cách lập bảng biểu cuối tháng trên Excel cho dân kế toán](#)


3. Hàm SUMIF
------------




* Hàm này dùng để tính tổng các ô trong vùng cần tính thỏa mãn một điều kiện đưa vào.

* Câu lệnh: = SUMIF(range, criteria,sum\_range), nghĩa là SUMIF (Vùng chứa điều kiện, điều kiện, vùng cần tính tổng)

* Ví dụ: Bài toán sử dụng hàm SUMIF



![vi-du-ham-sumif](https://hddtvn.com/wp-content/uploads/2021/01/vi-du-ham-sumif.jpg)Ví dụ hàm SUMIF trong excel
Nhìn vào ví dụ trên, ta thấy sau khi thực hiện câu lệnh hàm SUMIF, ta thu được kết quả tổng phụ cấp CTV là 500.000. Để kiểm tra lại, ta thấy có tất cả 5 CTV, với mỗi CTV được nhận phụ cấp là 100.000. Như vậy, kết quả tính trên là chính xác.


[**>>> Cách sử dụng hàm SUMIF trong Excel**](#)


4. Hàm COUNTIF
--------------




* Hàm dùng để đếm số ô thỏa mãn điều kiện trong một phạm vi nhất định.

* Câu lệnh: = COUNTIF (range,criteria), nghĩa là COUNTIF(phạm vi dữ liệu muốn đếm, điều kiện để một ô được đếm).

* Ví dụ: =COUNTIF(C5:C12,”<50”), nghĩa là đếm tất cả các ô thuộc dãy từ C5:C12 có chứa số nhỏ hơn 50).



[**>>> Mách bạn cách sử dụng hàm COUNTIF trong Excel**](#)


5. Hàm AND và OR
----------------


### Hàm AND




* Hàm này là phép VÀ, bao gồm các đối số là các hằng hoặc biểu thức logic, và chỉ đúng khi tất cả các đối số có giá trị đúng. Hàm trả về giá trị TRUE (1) khi tất cả các đối số của nó là đúng, trả về giá trị FALSE (0) nếu tồn tại ít nhất một đối số của nó là sai.

* Cú pháp: = AND(Logical1, Logical2,…) nghĩa là AND(đối 1, đối 2,..). Trong đó: Logical1, Logical2…là các biểu thức điều kiện.

* Ví dụ: =AND(C3>5,C3<100)



*Chú ý:*




* Các đối số phải mang giá trị logic hoặc mảng hay tham chiếu có mang giá trị logic.

* Nếu đối số là giá trị text hoặc Null (rỗng) thì bỏ qua những giá trị đó.

* Nếu vùng tham chiếu không chứa giá trị logic thì hàm trả về lỗi #VALUE!



![vi-du-ham-and](https://hddtvn.com/wp-content/uploads/2021/01/vi-du-ham-and.png)Ví dụ hàm AND trong excel
[**>>>Mẹo dùng hàm AND và kết hợp với các hàm logic khác trong Excel**](#)


### Hàm OR




* Hàm này là phép HOẶC, cũng bao gồm các đối số như hàm AND, nhưng chỉ sai khi tất cả các đối số có giá trị sai. Hàm trả về giá trị TRUE (1) nếu bất cứ một đối số nào của nó là đúng, trả về giá trị FALSE (0) nếu tất cả các đối số của nó là sai.

* Ví dụ: =OR(E7>03/02/76,E7>01/01/2015).



**>>>[Cách sử dụng hàm OR và cách kết hợp với các hàm logic khác trong Excel](#)**


6. Hàm MAX và MIN
-----------------


### Hàm MAX




* Trả về kết quả lớn nhất trong dãy được chọn, thường dùng để tìm số lớn nhất trong dãy.

* Câu lệnh: =MAX(Number 1, Number 2,…). Trong đó, Number1, Number2,…là dãy mà bạn muốn tìm giá trị lớn nhất.

* Ví dụ: =MAX(77,88,99) =99.



[**>>> Cách tìm giá trị lớn nhất và nhỏ nhất bằng hàm MAX và MIN trong Excel**](#)


### Hàm MIN




* Ngược lại với hàm MAX, hàm MIN trả về kết quả nhỏ nhất trong dãy được chọn, thường dùng để tìm số nhỏ nhất trong dãy.

* Câu lệnh: =MIN(Number 1, Number 2,…).Trong đó, Number1, Number2,…là dãy mà bạn muốn tìm giá trị nhỏ nhất.

* Ví dụ: =MIN(77,88,99) =77.



![vi-du-ham-min-max](https://hddtvn.com/wp-content/uploads/2021/01/vi-du-ham-min-max.png)Ví dụ hàm MIN và MAX trong excel
[**>>> Cách tìm giá trị lớn nhất và nhỏ nhất bằng hàm MAX và MIN trong Excel**](#)


7. Hàm IF
---------




* Hàm trả về kết quả 1 nếu điều kiện đúng, hàm trả về kết quả 2 nếu điều kiện sai.

* Câu lệnh: = IF(Logical\_test,[value\_if\_true],[value\_if\_false]).



*Trong đó:*


**Logical\_test** là biểu thức điều kiện.


**[value\_if\_true]** là giá trị được trả về nếu điều kiện đúng.


**[value\_if\_false]** là giá trị được trả về nếu điều kiện sai.


**[>>>Cách dùng hàm IF và cách kết hợp với các hàm khác trong Excel](#)**


Trên đây là một số hàm excel sử dụng phổ biến nhất trong nghiệp vụ kế toán của doanh nghiệp. Bên cạnh những hàm trên, excel còn rất nhiều hàm khác với nhiều chức năng khác nhau. Tuy nhiên, trước khi nắm được những hàm excel nâng cao, bạn nên tập trung nắm vững được những hàm excel cơ bản nhất như bài viết đã liệt kê ở trên.


Bằng cách làm như vậy, bạn sẽ tự tin giải quyết được các vấn đề liên quan đến nghiệp vụ kế toán và có cơ hội phát triển hơn trong nghề nghiệp này.



moreExcel được coi là một phần không thể thiếu trong kế toán doanh nghiệp. Việc…

