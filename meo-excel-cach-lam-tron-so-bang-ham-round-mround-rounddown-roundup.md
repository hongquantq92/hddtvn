Mẹo Excel: Cách làm tròn số bằng hàm ROUND, MROUND, ROUNDDOWN, ROUNDUP
======================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/meo-excel-cach-lam-tron-so-bang-ham-round-mround-rounddown-roundup/)Trong quá trình làm việc trên Excel, bạn sẽ rất hay gặp trường hợp phải làm tròn số. Tùy theo nhu cầu mà bạn cần phải sử dụng những hàm khác nhau để làm tròn. Bài viết dưới đây sẽ hướng dẫn các bạn sử dụng hàm ROUND, MROUND, ROUNDDOWN, ROUNDUP để làm tròn số …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Trong quá trình làm việc trên Excel, bạn sẽ rất hay gặp trường hợp phải làm tròn số. Tùy theo nhu cầu mà bạn cần phải sử dụng những hàm khác nhau để làm tròn. Bài viết dưới đây sẽ hướng dẫn các bạn sử dụng hàm ROUND, MROUND, ROUNDDOWN, ROUNDUP để làm tròn số trong các trường hợp khác nhau.**


### 1. Làm tròn số trong Excel bằng hàm ROUND


Cấu trúc của hàm ROUND: **=ROUND(number; num\_ditgits)**


*Trong đó:*




* **number**: Số tiền được làm tròn, cần xét làm tròn

* **num\_ditgits**: Phần được làm tròn.



#### a. Làm tròn hết phần thập phân (làm tròn tới 0)


Quy tắc làm tròn hết phần thập phân như sau:




* Nếu phần xét làm tròn nhỏ hơn 0,5 thì sẽ được làm tròn xuống bằng 0

* Nếu phần xét làm tròn lớn hơn hoặc bằng 0,5 thì sẽ được làm tròn lên 1



Công thức làm tròn hết phần thập phân như sau: **=ROUND(number;0)**


Ví dụ ta cần làm tròn hết phần thập phân của cột E (SẢN LƯỢNG)


[![](https://hddtvn.com/wp-content/uploads/2021/01/fR6xFb9.png)](https://hddtvn.com/wp-content/uploads/2021/01/fR6xFb9.png)


Trong hình trên ta có thể thấy toàn bộ phần thập phân đã được làm tròn hết tại cột F:




* Số thứ 1 với phần thập phân là 0,068 được làm tròn về 0

* Số thứ 3 với phần thập phân 0,624 được làm tròn thêm 1 để từ 67.254 thành 67.255



#### b. Làm tròn tới phần nghìn


Làm tròn tới phần nghìn là một yêu cầu khá hay gặp khi làm kế toán. Quy tắc làm tròn tới phần nghìn như sau:




* Nếu số tiền lẻ dưới 500 thì làm tròn xuống 0

* Nếu số tiền lẻ lớn hơn hoặc bằng 500 thì làm tròn lên thêm 1000



Công thức làm tròn tới phần nghìn như sau: **=ROUND(number;-3)**


Ví dụ ta cần làm tròn tới phần nghìn của cột E (SẢN LƯỢNG):


![](https://hddtvn.com/wp-content/uploads/2021/01/f77gzPp.png)


Trong hình trên ta có thể thấy toàn bộ phần nghìn đã được làm tròn hết tại cột F




* Số thứ 1 là 1.054,068 được làm tròn phần 054,068 thành 000,000 ra kết quả là 1.000,000

* Số thứ 5 là 49.951,399 được làm tròn phần 951,399 thành 1.000,000 ra kết quả là 50.000,000



Việc làm tròn này thường áp dụng với số tiền lớn hàng trăm triệu, hàng tỉ. Khi đó phần làm tròn chỉ mang giá trị rất nhỏ nên ít ảnh hưởng tới việc tính toán trong kế toán.


### 2. Làm tròn đến bội số của một số khác bằng hàm MROUND


Cú pháp: **=MROUND(number; multiple)**


Trong đó:




* **number**: Giá trị muốn làm tròn, là tham số bắt buộc.

* **multiple**: Bội số muốn làm tròn tới, là tham số bắt buộc.



*Chú ý:*




* Hàm MROUND làm tròn lên hướng ra xa số 0 khi và chỉ khi số dư của phép chia number cho multiple lớn hơn hoặc bằng một nửa giá trị multiple.

* Trường hợp giá trị number và multiple khác dấu -> hàm trả về giá trị lỗi #NUM!



*Ví dụ:*  

MROUND(5, 2) = 6 (do 5/2 > 2/2, bội số của 2 gần nhất mà lớn hơn 5 là 6)  

MROUND(11, 5) = 10 (do 11/5 < 5/2, bội số của 5 gần nhất mà nhỏ hơn 11 là 10)  

MROUND(13, 5) = 15 (do 13/5 > 5/2, bội số của 5 gần nhất mà lớn hơn 13 là 15)  

MROUND(5, 5) = 5 (number và multiple bằng nhau)  

MROUND(7.31, 0.5) = 7.5 (do 7.31/0.5 > 0.5/2, bội số của 0.5 gần nhất mà lớn hơn 7.31 là 7.5)  

MROUND(-11, -5) = -10 (do -11/-5 > -5/2, bội số của -5 gần nhất mà lớn hơn -11 là -10)  

MROUND(-11, 5) = #NUM! (number và multiple khác dấu)


### 3. Làm tròn số bằng hàm ROUNDDOWN và ROUNDUP


Hai hàm này, về cơ bản thì khá giống hàm ROUND. Chỉ khác là chúng chỉ làm tròn theo một chiều:




* Hàm ROUNDDOWN làm tròn số xuống (tiến gần tới 0 hay số đằng trước nó)

* Hàm ROUNDUP làm tròn một số lên (tiến gần đến số đằng sau nó hay làm tròn hướng ra xa số 0)



*Cú pháp:*


**= ROUNDDOWN(number;num\_digits)**


**= ROUNDUP(number;num\_digits)**


**num\_digits**: Là một số nguyên, chỉ cách mà bạn muốn làm tròn  

**num\_digits** > 0 : làm tròn đến số thập phân được chỉ định  

**num\_digits** = 0 : làm tròn đến số nguyên gần nhất  

**num\_digits** < 0 : làm tròn đến phần nguyên được chỉ định


Ví dụ: So sánh giữa hàm **ROUNDDOWN** và hàm **ROUNDUP**


![](https://hddtvn.com/wp-content/uploads/2021/01/uOioY8f.png)


*Như vậy bài viết trên đã hướng dẫn các bạn cách làm tròn số bằng hàm ROUND, MROUND, ROUNDDOWN, ROUNDUP. Chúc các bạn thành công!*


moreTrong quá trình làm việc trên Excel, bạn sẽ rất hay gặp trường hợp phải…

