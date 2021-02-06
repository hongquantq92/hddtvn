Cách tách họ và tên thành hai cột riêng biệt trong Excel
========================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-tach-ho-va-ten-thanh-hai-cot-rieng-biet-trong-excel/)Bạn đã nhập họ và tên đầy đủ của mọi người trong cùng một ô. Nhưng sau đó lại cần tách riêng họ tên đệm và tên ra hai ô khác nhau để thuận tiện cho việc quản lý. Làm sao để tách cột Họ và Tên thành 2 cột Họ và cột Tên riêng …
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Bạn đã nhập họ và tên đầy đủ của mọi người trong cùng một ô. Nhưng sau đó lại cần tách riêng họ tên đệm và tên ra hai ô khác nhau để thuận tiện cho việc quản lý. Làm sao để tách cột Họ và Tên thành 2 cột Họ và cột Tên riêng biệt? Việc tưởng chừng như đơn giản nhưng lại không nhiều người chưa biết cách làm. Hãy theo dõi bài viết sau để nắm rõ cách thực hiện nhé.**


Ví dụ ta có danh sách tên như hình dưới. Yêu cầu cần tách riêng họ và tên của từng người ra. Để thực hiện, các bạn hãy làm theo các bước sau đây nhé.


![](https://hddtvn.com/wp-content/uploads/2021/01/WGi3XkV.png)


### Cách 1: Sử dụng Replace và hàm


**Bước 1:** Đầu tiên, các bạn cần sao chép toàn bộ họ và tên của tất cả mọi người sang cột **TÊN**.


![](https://hddtvn.com/wp-content/uploads/2021/01/SSCHTSP.png)


**Bước 2:** Tiếp theo, các bạn bôi đen toàn bộ cột **TÊN** rồi nhấn **Ctrl + H**. Hộp thoại **Find and Replace** hiện ra thì các bạn nhập dấu ***** và dấu cách vào mục **Find what** rồi nhấn **Replace All**. Chỉ cần như vậy là tất cả phần họ sẽ bị loại bỏ và chỉ giữ lại phần tên.


![](https://hddtvn.com/wp-content/uploads/2021/01/UFxXnOD.png)


**Bước 3:** Tiếp theo, để tách được phần họ. Các bạn nhập công thức sau và ô **C2**:


**=LEFT(B2;LEN(B2)-LEN(D2))**


*Trong đó:*




* **LEFT** là hàm cắt chuỗi từ trái qua, cú pháp =LEFT(ô chứa chuỗi cần cắt từ trái qua, số ký tự sẽ cắt)

* **LEN** là hàm lấy độ dài của chuỗi, cú pháp =LEN(ô chứa chuỗi cần lấy độ dài)



![](https://hddtvn.com/wp-content/uploads/2021/01/BNIgzPU.png)


**Bước 4:** Sao chép công thức trên cho tất cả ô còn lại ở cột **HỌ** là ta sẽ thu được phần họ đã được tách riêng ra một cách nhanh chóng.


![](https://hddtvn.com/wp-content/uploads/2021/01/J5rs2Fg.png)


### Cách 2: Sử dụng nhiều hàm phức tạp


**Bước 1:** Đầu tiên, các bạn tách phần tên bằng cách nhập công thức sau vào ô **D2**:


**=RIGHT(B2;LEN(B2)-FIND("@";SUBSTITUTE(B2;" ";"@";LEN(B2)-LEN(SUBSTITUTE(B2;" ";"")))))**


![](https://hddtvn.com/wp-content/uploads/2021/01/aO1CyAT.png)


**Bước 2:** Sau đó các bạn sao chép công thức trên cho các ô còn lại tại cột **TÊN**. Kết quả ta sẽ thu được phần tên đã được tách ra một cách nhanh chóng.


![](https://hddtvn.com/wp-content/uploads/2021/01/o4zRPLQ.png)


**Bước 3:** Tiếp theo, để tách được phần họ. Các bạn nhập công thức sau và ô **C2**:


**=LEFT(B2;LEN(B2)-LEN(D2))**


*Trong đó:*




* **LEFT** là hàm cắt chuỗi từ trái qua, cú pháp =LEFT(ô chứa chuỗi cần cắt từ trái qua, số ký tự sẽ cắt)

* **LEN** là hàm lấy độ dài của chuỗi, cú pháp =LEN(ô chứa chuỗi cần lấy độ dài)



![](https://hddtvn.com/wp-content/uploads/2021/01/BNIgzPU.png)


**Bước 4:** Sao chép công thức trên cho tất cả ô còn lại ở cột **HỌ** là ta sẽ thu được phần họ đã được tách riêng ra một cách nhanh chóng.


[![](https://hddtvn.com/wp-content/uploads/2021/01/J5rs2Fg.png)](https://hddtvn.com/wp-content/uploads/2021/01/J5rs2Fg.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách tách riêng phần họ và tên trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreBạn đã nhập họ và tên đầy đủ của mọi người trong cùng một ô….

