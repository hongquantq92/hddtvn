Cách sử dụng hàm DATE và cách kết hợp với các hàm khác trong Excel
==================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-date-va-cach-ket-hop-voi-cac-ham-khac-trong-excel/)Excel không lưu trữ ngày, tháng dưới định dạng ngày. Thay vào đó, Excel lưu trữ ngày như một dãy số và đây là chính là nguyên ngân chính gây nhầm lẫn. Do đó khi cần tính toán ngày trong Excel, DATE là hàm cơ bản nhất. Hàm DATE được sử dụng để tạo thành …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Excel không lưu trữ ngày, tháng dưới định dạng ngày. Thay vào đó, Excel lưu trữ ngày như một dãy số và đây là chính là nguyên ngân chính gây nhầm lẫn. Do đó khi cần tính toán ngày trong Excel, DATE là hàm cơ bản nhất. Hàm DATE được sử dụng để tạo thành một ngày từ ba giá trị ngày, tháng, năm riêng biệt. Hãy theo dõi bài viết sau đây để biết cách sử dụng hàm DATE trong Excel.**


### 1. Cấu trúc hàm DATE


Cú pháp hàm: **=DATE(year; month; day)**


*Trong đó:*




* **year**: đối số bắt buộc, là năm sử dụng khi tạo ngày. Đối số **year** có thể bao gồm từ một đến bốn chữ số. Theo mặc định, Excel sử dụng hệ thống ngày tháng 1900, tức là ngày đầu tiên là 01/01/1900.

* **month**: đối số bắt buộc, là tháng sử dụng khi tạo ngày. Phải là số nguyên dương hoặc âm.

* **day**: đối số bắt buộc, là ngày sử dụng khi tạo ngày. Phải là số nguyên dương hoặc âm.



*Lưu ý:*




* Nếu **year** từ 0 đến 1899 (gồm cả 0 và 1899), hàm sẽ cộng thêm 1900 vào giá trị đó để tính năm.

* Nếu **year** từ 1900 đến 9999 (gồm cả 1990 và 999), hàm sẽ dùng giá trị đó làm năm.

* Nếu **year** nhỏ hơn 0 hoặc lớn hơn 10000, hàm sẽ trả về giá trị lỗi #NUM!.

* Nếu **month** lớn hơn 12, hàm sẽ thêm số tháng lớn hơn vào tháng đầu tiên trong năm tiếp theo.

* Nếu **month** âm, hàm sẽ trừ số tháng ngược lại vào năm trước.

* Nếu **day** lớn hơn số ngày trong tháng đã xác định, hàm sẽ thêm số ngày lớn đó vào tháng tiếp theo.

* Nếu **day** nhỏ hơn 1,hàm sẽ trừ số ngày lớn hơn đó vào tháng trước.



### 2. Cách sử dụng hàm DATE


#### a. Sử dụng hàm DATE để gộp ngày, tháng, năm


Ví dụ ta cần ghép các ngày, tháng, năm sau đây vào thành 1:


![](https://hddtvn.com/wp-content/uploads/2021/01/S21excY.png)


Áp dụng cấu trúc hàm như trên ta có công thức cho ô D2 như sau: **=DATE(C2;B2;A2)**


Hoặc ta cũng có thể viết thẳng số ngày, tháng, năm vào hàm như sau: **=DATE(20;02;2020)**


Sao chép công thức trên cho các ô bên dưới ta thu được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/VUTjiDh.png)


Như kết quả thu được ta có thể thấy:




* Tại ô D3: do năm tại ô C3 là 82 < 1900 nên hàm đã tự động cộng thêm 1900 vào để thành số năm là 1982

* Tại ô D4: do năm là 11000 > 10000 nên hàm trả về giá trị lỗi #NUM!

* Tại ô D5: do tháng là 15 > 12 nên hàm đã cộng thêm số tháng lớn hơn vào năm tiếp theo của 2018 là 2019

* Tại ô D6: do tháng là -4 là số âm nên hàm đã trừ ngược số tháng lại năm trước 2009 là 2008

* Tại ô D7: do ngày là 33 lớn hơn số ngày trong tháng 05 nên hàm đã cộng số ngày lớn hơn vào tháng tiếp theo

* Tại ô D8: do số ngày là -3 là số âm nên hàm đã trừ ngược số ngày lại tháng trước đó.



#### b. Sử dụng hàm DATE kết hợp với các hàm DAY, MONTH, YEAR


Ví dụ ta có ngày ban đầu, sau đó cần tính 4 ngày sau ngày đó, 5 tháng sau ngày đó, 3 năm sau ngày đó.


![](https://hddtvn.com/wp-content/uploads/2021/01/UWu8Esb.png)


Bằng cách kết hợp với các hàm YEAR, MONTH, DAY. Ta có công thức tính 4 ngày sau ngày ban đầu cho ô B2 như sau:


**=DATE(YEAR(A2);MONTH(A2);DAY(A2)+4)**


Copy công thức cho các ô còn lại trong cột B ta thu được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/fGtahqs.png)


Tương tự để tính 5 tháng sau ngày ban đầu ta có công thức cho ô C2:


**=DATE(YEAR(A2);MONTH(A2)+5;DAY(A2))**


![](https://hddtvn.com/wp-content/uploads/2021/01/wfZJAZ9.png)


Công thức để tính 3 năm sau ngày ban đầu cho ô D2:


**=DATE(YEAR(A2)+3;MONTH(A2);DAY(A2))**


![](https://hddtvn.com/wp-content/uploads/2021/01/ytT8vDG.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm DATE trong Excel. Chúc các bạn thành công!*


moreExcel không lưu trữ ngày, tháng dưới định dạng ngày. Thay vào đó, Excel lưu…

