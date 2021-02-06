Cách sử dụng hàm DCOUNTA để đếm ô thỏa mãn điều kiện
====================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-dcounta-de-dem-o-thoa-man-dieu-kien/)DCOUNTA là hàm thống kê được sử dụng khá phổ biến trong Excel. Hàm có tác dụng đếm những ô thỏa mãn điều kiện được cho trước trong Excel. Ví dụ bạn có một list các sản phẩm và muốn tìm các mặt hàng có đơn giá, được giảm giá, tăng giá hoặc có thay …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**DCOUNTA là hàm thống kê được sử dụng khá phổ biến trong Excel. Hàm có tác dụng đếm những ô thỏa mãn điều kiện được cho trước trong Excel. Ví dụ bạn có một list các sản phẩm và muốn tìm các mặt hàng có đơn giá, được giảm giá, tăng giá hoặc có thay đổi về giá… thì hãy sử dụng tới hàm DCOUNTA. Mời****bạn theo dõi bài viết sau để nắm rõ cách sử dụng hàm DCOUNTA nhé.**


### 1. Cấu trúc hàm DCOUNTA


Cú pháp hàm: **=DCOUNTA(Database; Field; Criteria)**


*Trong đó:*




* **Database**: đối số bắt buộc, là phạm vi ô hoặc danh sách hay cơ sở dữ liệu, hàng đầu tiên của danh sách phải chứa tiêu đề cho mỗi cột.

* **Field**: đối số bắt buộc, là cột được dùng trong hàm, hay tiêu đề cột nếu nhập trực tiếp tiêu đề cột đặt trong dấu ngoặc kép, hoặc có thể nhập số chỉ số thứ tự của cột với cột đầu tiên là cột số 1.

* **Criteria**: đối số bắt buộc, là phạm vi ô chứa điều kiện. Có thể dùng bất kì phạm vi nào nhưng chứa ít nhất 1 tiêu đề cột và 1 ô chứa điều kiện dưới ô tiêu đề cột.



*Lưu ý:*




* Có thể dùng bất kì phạm vi nào cho đối số criteria nhưng phải chứa ít nhất 1 tiêu đề cột và 1 ô chứa điều kiện dưới ô tiêu đề cột.

* Không đặt phạm vi điều kiện bên dưới danh sách vì thông tin mới sẽ được thêm vào cuối danh sách.

* Phạm vi điều kiện không ghi đè lên danh sách.

* Để thao tác toàn bộ một cột trong danh sách nên nhập 1 dòng trống bên dưới hay bên trên tiêu đề cột trong phạm vi điều kiện.



### 2. Cách sử dụng hàm DCOUNTA trong Excel


Ví dụ ta có bảng dữ liệu như hình dưới. Yêu cầu đếm những số lượng mặt hàng thảo mãn điều kiện sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/0GbbsJe.png)


#### a. Đếm số lượng mặt hàng có đơn giá thay đổi


Áp dụng cấu trúc hàm bên trên, ta có công thức đếm số lượng mặt hàng có đơn giá thay đổi như sau:


**=DCOUNTA(A1:F10;F1;I1:I2)**


Kết quả ta sẽ thu được số lượng mặt hàng có đơn giá thay đổi là 6 mặt hàng.


![](https://hddtvn.com/wp-content/uploads/2021/01/tidhhXk.png)


#### b. Đếm số lượng mặt hàng có đơn giá tăng


Để đếm số lượng mặt hàng có đơn giá tăng, các bạn cần nhập **Tăng *** vào điều kiện. Sau đó ta có công thức đếm như sau:


**=DCOUNTA(A1:F10;F1;I1:I2)**


Kết quả ta sẽ thu được số lượng mặt hàng có đơn giá tăng là 2 mặt hàng.


![](https://hddtvn.com/wp-content/uploads/2021/01/aUfechR.png)


#### c. Đếm số lượng mặt hàng có đơn giá giảm trên 10%


Để đếm số lượng mặt hàng có đơn giá giảm trên 10%, các bạn cần nhập **Giảm 1*%** vào điều kiện. Sau đó ta có công thức đếm như sau:


**=DCOUNTA(A1:F10;F1;I1:I2)**


Kết quả ta sẽ thu được số lượng mặt hàng có đơn giá giảm trên 10% là 3 mặt hàng.


[![](https://hddtvn.com/wp-content/uploads/2021/01/bclDDoj.png)](https://hddtvn.com/wp-content/uploads/2021/01/bclDDoj.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm DCOUNTA để đếm số lượng ô thỏa mãn điều kiện. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreDCOUNTA là hàm thống kê được sử dụng khá phổ biến trong Excel. Hàm có…

