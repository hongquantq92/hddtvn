Mẹo đếm số lượng ô trống và tô màu ô trống trong Excel
======================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/meo-dem-so-luong-o-trong-va-to-mau-o-trong-trong-excel/)Trong một bảng Excel không phải lúc nào các ô cũng đầy đủ số liệu, sẽ có những ô chứa giá trị rỗng hay trống số liệu. Những ô trống ngẫu nhiên này ít nhiều ảnh hưởng tới việc tính toán giá trị. Trong Excel có một cách để xác định nhanh tổng số ô …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Trong một bảng Excel không phải lúc nào các ô cũng đầy đủ số liệu, sẽ có những ô chứa giá trị rỗng hay trống số liệu. Những ô trống ngẫu nhiên này ít nhiều ảnh hưởng tới việc tính toán giá trị. Trong Excel có một cách để xác định nhanh tổng số ô trống, ô không chứa số liệu và tô màu những ô không có giá trị đó. Thao tác này sẽ giúp bạn thuận tiện hơn trong việc tính toán. Trong bài viết này, hddtvn xin chia sẻ với các bạn mẹo đếm số lượng ô trống và tô màu những ô trống này nhé.**


![Mẹo đếm số lượng ô trống và tô màu ô trống trong Excel](https://hddtvn.com/wp-content/uploads/2021/01/IPV8z9y.png "Mẹo đếm số lượng ô trống và tô màu ô trống trong Excel")


Ví dụ ta có bảng dữ liệu như hình dưới. Như các bạn có thể thấy, trong bảng dưới có rất nhiều ô trống do nhập thiếu dữ liệu. Muốn biết có bao nhiêu ô trống trong bảng thì cần phải đếm nhưng nếu đếm thủ công thì sẽ rất dễ sai do khó nhìn. Để giải quyết vấn đề này thì ta có thể sử dụng hàm để đếm số lượng ô trống và tô màu những ô trống đó một cách nhanh chóng.


![](https://hddtvn.com/wp-content/uploads/2021/01/EQw9hrP.png)


### 1. Đếm số lượng ô trống


Để đếm được số lượng ô trống trong bảng dữ liệu trên, các bạn có thể sử dụng hàm COUNTBLANK. Cấu trúc của hàm COUNTBLANK như sau:


Cú pháp hàm: **=COUNTBLANK(range)**


Trong đó: **Range**là mảng mà bạn muốn đếm các ô trống.


Áp dụng công thức hàm vào trường hợp này ta có công thức đếm số lượng ô trống của bảng dữ liệu trên như sau:


**=COUNTBLANK(C1:E11)**


Như vậy là ta có thể thấy kết quả số lượng ô trống trong bảng trên là 9 ô.


![](https://hddtvn.com/wp-content/uploads/2021/01/WPiPtpP.png)


### 2. Tô màu ô trống


Để tô màu cho những ô trống trong bảng, đầu tiên các bạn cần bôi đen toàn bộ vùng dữ liệu có ô trống. Sau đó các bạn chọn mục **Conditional Formatting** trong thẻ **Home.** Thanh cuộn hiện ra các bạn chọn mục **New Rule**.


![](https://hddtvn.com/wp-content/uploads/2021/01/AndzKjg.png)


Lúc này, hộp thoại **New Formatting Rule** hiện ra. Các bạn chọn mục **Use a formula to determine which cells to format** để tạo công thức. Sau đó các bạn nhập công thức **=ISBLANK(C2)** vào mục **Format values where this formula is true**. Sau đó chọn mục **Format** để chọn màu tô.


![](https://hddtvn.com/wp-content/uploads/2021/01/nA2xcrf.png)


Lúc này, hộp thoại Format Cells hiện ra. Các bạn chọn thẻ **Fill**. Sau đó chọn màu tô tại mục **Background Color**. Cuối cùng nhấn **OK** để hoàn tất.


![](https://hddtvn.com/wp-content/uploads/2021/01/a7uwJco.png)


Chỉ cần như vậy là những ô trống trong bảng sẽ được tô màu để bạn có thể dễ dàng nhận ra chúng.


![](https://hddtvn.com/wp-content/uploads/2021/01/IPV8z9y-1.png)


Nếu bạn đã xử lý xong dữ liệu và không muốn tô màu ô trống nữa thì chọn **Conditional Formatting** => **Clear Rules** => **Clear Rules from Selected Cells**.


![](https://hddtvn.com/wp-content/uploads/2021/01/CkVrA3e.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách đếm số lượng ô trống cũng như cách tô màu ô trống trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreTrong một bảng Excel không phải lúc nào các ô cũng đầy đủ số liệu,…

