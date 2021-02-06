Cách sử dụng hàm Match và cách kết hợp với các hàm tìm kiếm Lookup
==================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-match-va-cach-ket-hop-voi-cac-ham-tim-kiem-lookup/)Hàm Match là hàm giúp tìm kiếm một mục cụ thể trong một dãy các ô, và đưa ra vị trí tương đối của mục đó. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm Match và kết hợp hàm Match với hàm Vlookup, Hlookup. 1. Cú pháp hàm và cách …
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm Match là hàm giúp tìm kiếm một mục cụ thể trong một dãy các ô, và đưa ra vị trí tương đối của mục đó. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm Match và kết hợp hàm Match với hàm Vlookup, Hlookup.**


![](https://hddtvn.com/wp-content/uploads/2021/01/Co-che-hoat-dong-vlookup-hlookup-01.png)


### 1. Cú pháp hàm và cách sử dụng


Cú pháp của hàm Match: **MATCH(lookup\_value;lookup\_array;[match\_type])**


*Trong đó:*


**Lookup\_value**: là giá trị cần tìm. Có thể là dạng số, chữ hoặc ô tham chiếu


**Lookup\_array**: là dãy các ô cần tìm kiếm


**Match\_type**: xác định loại tìm kiếm. Tham số không bắt buộc. Có thể nhận một trong các giá trị sau: 1, 0, -1 Nếu bạn không điền giá trị, Excel sẽ mặc định kiểu khớp =1.




* 1 hoặc không ghi gì: tìm giá trị lớn nhất trong dãy tìm kiếm sao cho giá trị đó nhỏ hơn hoặc bằng giá trị tìm kiếm. Loại này yêu cầu sắp xếp dãy tìm kiếm theo thứ tự tăng dần, từ nhỏ nhất đến lớn nhất hoặc từ A đến Z

* 0: tìm giá trị đầu tiên trong dãy bằng đúng giá trị tìm kiếm. Không yêu cầu sắp xếp dãy tìm kiếm

* -1: tìm giá trị nhỏ nhất trong dãy lớn hơn hoặc bằng giá trị tìm kiếm. Dãy tìm kiếm nên được sắp xếp theo thứ tự giảm dần, từ to nhất đến nhỏ nhất hoặc từ Z đến A.’



#### *Lưu ý:*




* Hàm MATCH chỉ trả lại vị trí tương đối của giá trị cần tìm kiếm trong một dãy, không phải giá trị của chính nó  

Hàm Match không phân biệt chữ in hoa và chữ thường.

* Nếu chuỗi tìm kiếm chứa một vài giá trị tìm kiếm, vị trí của giá trị đầu tiên sẽ được trả về  

Nếu không tìm thấy giá trị phù hợp, kết quả sẽ trả về giá trị lỗi #N/A.

* Nếu bạn dùng match\_type =1 , Mảng tìm kiếm phải được sắp xếp theo thứ tự tăng dần.

* Nếu bạn dùng match\_type = -1 , Mảng tìm kiếm phải được sắp xếp theo thứ tự giảm dần.

* Nếu lookup\_value là dạng chữ thì phải cho vào dấu ngoặc kép “”.

* Nếu match\_type = 0 và lookup\_value là chuỗi văn bản, thì bạn có thể dùng ký tự dấu hỏi (?) để đại diện cho bất kì kí tự đơn nào và dấu sao (*) phù hợp với bất kỳ chuỗi ký tự nào. Nếu bạn muốn tìm một dấu chấm hỏi hay dấu sao thực, hãy gõ một dấu ngã (~) trước ký tự đó.

* Nếu trong mảng tìm kiếm lookup\_array có nhiều hơn một giá trị của lookup\_value, thì kết quả sẽ trả về kết quả nhỏ nhất.



Ví dụ: Tên hàng ở cột B và Doanh thu các năm ở các cột C, D, E. Để tìm ra vị trí của loại hàng (ví dụ Bia lon), sử dụng công thức sau: **=MATCH(B11;B2:B7;0)**


[![](https://hddtvn.com/wp-content/uploads/2021/01/OKeXXyW.png)](https://hddtvn.com/wp-content/uploads/2021/01/OKeXXyW.png)


Công thức MATCH cho chúng ta biết rằng Bia long đứng ở vị trí thứ 3 trong dãy giá trị tìm kiếm.


### 2. Kết hợp hàm Match với các hàm tìm kiếm Lookup


Bạn có thể dùng hàm Match để lấy vị trí tương đối của cột/dòng mà bạn cần trả lại, và cung cấp số cột/dòng cho tham số Row­\_index\_number đối với hàm Hlookup / tham số Col­\_index\_number đối với hàm Vlookup.


#### a. Hàm Match kết hợp với hàm Vlookup


Hàm MATCH được dùng để xác định vị trí tương đối của giá trị tìm kiếm, do đó hoàn toàn phù hợp đối với **col\_index\_num** của hàm VLOOKUP. Do đó thay vì chỉ rõ cột trả lại như một số không thay đổi, bạn sử dụng hàm MATCH để biết vị trí hiện tại của cột đó.


Ví dụ ta cần tìm Doanh thu 2018 của Bia lon. Thông thường ta sẽ có công thức sau:


**=VLOOKUP(A11;A1:D7;3;FALSE)**


Tham số **col\_index\_num** được đặt là 3 vì Doanh thu 2018 mà chúng ta muốn kéo nằm ở cột thứ 3 trong bảng.


![](https://hddtvn.com/wp-content/uploads/2021/01/kyZWUmt.png)


Nhưng nếu bạn chèn thêm cột hoặc xóa bớt cột trong bảng đi. Thì kết quả sẽ không còn đúng nữa. Do cột Doanh thu 2018 không còn là cột thứ 3 của bảng.


![](https://hddtvn.com/wp-content/uploads/2021/01/imdMh3H.png)


Để giải quyết điều này, ta chỉ cần thêm hàm Match kết hợp vào hàm Vlookup


**=MATCH(B10;A1:D1;0)**


*Trong đó:*




* **B10**: là tên chính xác của cột cần tìm kiếm

* **A1:D1**: là dãy tìm kiếm chứa bảng



Sau đó gộp hàm Match vào tham số col\_index\_num của hàm Vlookup như sau:


**=VLOOKUP(A11;A1:D7;MATCH(B10;A1:D1;0);FALSE)**


![](https://hddtvn.com/wp-content/uploads/2021/01/YCquxZv.png)


Lúc này hàm của chúng ta có thể tự điều chỉnh tham chiếu dù có chèn thêm hay xóa bớt bao nhiêu cột đi nữa.


#### b. Hàm Match kết hợp với hàm Hlookup


Tương tự như hàm Vlookup, ta cũng có thể sử dụng hàm Match để kết hợp với hàm HLOOKUP. Ta sử dụng hàm Match để lấy vị trí tương đối của cột cần trả lại, và cung cấp số của cột đó cho tham số **row\_index\_number** của hàm Hlookup.


Công thức hàm Hlookup trong trường hợp này sẽ như sau:


**=HLOOKUP(B10;A1:D7;MATCH(A11;A1:A7;0);FALSE)**


![](https://hddtvn.com/wp-content/uploads/2021/01/iL1iyzS.png)


*Như vậy, bài viết trên đã giới thiệu với các bạn cách sử dụng hàm Match và cách kết hợp hàm Match với các hàm tìm kiếm Vlookup, Hlookup. Chúc các bạn thành công!*



moreHàm Match là hàm giúp tìm kiếm một mục cụ thể trong một dãy các…

