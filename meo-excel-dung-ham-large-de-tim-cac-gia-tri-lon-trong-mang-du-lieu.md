Mẹo Excel: Dùng hàm LARGE để tìm các giá trị lớn trong mảng dữ liệu
===================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/meo-excel-dung-ham-large-de-tim-cac-gia-tri-lon-trong-mang-du-lieu/)Trong trường hợp muốn tìm theo giá trị lớn nhất, chúng ta có thể sử dụng hàm MAX. Nhưng với những giá trị lớn thứ hai, thứ ba hoặc thứ n của bảng dữ liệu thì sao? Lúc này chúng ta cần tới hàm LARGE. Bài viết sau đây sẽ hướng dẫn các bạn cách …
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Trong trường hợp muốn tìm theo giá trị lớn nhất, chúng ta có thể sử dụng hàm MAX. Nhưng với những giá trị lớn thứ hai, thứ ba hoặc thứ n của bảng dữ liệu thì sao? Lúc này chúng ta cần tới hàm LARGE. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm LARGE trong Excel.**


### 1. Cấu trúc hàm LARGE


Cú pháp hàm: **=LARGE(array; k)**


*Trong đó:*




* **Array**: đối số bắt buộc, là mảng hoặc phạm vi dữ liệu mà bạn muốn xác định giá trị lớn thứ k trong đó.

* **K**: đối số bắt buộc, là vị trí (tính từ lớn nhất) trong mảng dữ liệu cần trả về.



*Lưu ý:*




* Nếu n là số điểm dữ liệu trong mảng, thì hàm LARGE(mảng,1) trả về giá trị lớn nhất, và hàm LARGE(mảng,n) trả về giá trị nhỏ nhất.

* Nếu đối số mảng dữ liệu để trống, hàm LARGE trả về giá trị lỗi #NUM! .

* Nếu k ≤ 0 hoặc nếu nó lớn hơn số điểm dữ liệu trong mảng, hàm LARGE trả về giá trị lỗi #NUM!.



### 2. Cách sử dụng hàm LARGE


Vị dụ ta có bảng dữ liệu doanh thu các mặt hàng nước giải khát như hình dưới. Yêu cầu tìm mặt hàng có doanh thu lớn thứ 3


![](https://hddtvn.com/wp-content/uploads/2021/01/5stx9c5.png)


Để tìm mặt hàng có doanh thu lớn thứ 3, chúng ta phải qua 2 bước:




* Bước 1: Tìm giá trị doanh thu lớn thứ 3 với hàm LARGE

* Bước 2: Kết hợp hàm LARGE với hàm INDEX + MATCH để tìm tên mặt hàng tương ứng theo giá trị mức doanh thu lớn thứ 3



#### Bước 1: Sử dụng hàm LARGE để tìm doanh thu lớn thứ 3


Áp dụng cúp pháp hàm bên trên ta có công thức tìm doanh thu lớn thứ 3 như sau:


**=LARGE(C2:C10;3)**


![](https://hddtvn.com/wp-content/uploads/2021/01/HhkP7cL.png)


#### Bước 2: Kết hợp hàm LARGE với hàm INDEX + MATCH để tìm tên mặt hàng


Ta có công thức tìm tên mặt hàng như sau:


**=INDEX(B2:B10;MATCH(C12;C2:C10;0))**


![](https://hddtvn.com/wp-content/uploads/2021/01/qxOMKJf.png)


Hoặc ta cũng có thể viết thẳng hàm LARGE vào công thức:


**=INDEX(B2:B10;MATCH(LARGE(C2:C10;3);C2:C10;0))**


![](https://hddtvn.com/wp-content/uploads/2021/01/dgvMVMf.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm LARGE trong Excel. Chúc các bạn thành công!*


moreTrong trường hợp muốn tìm theo giá trị lớn nhất, chúng ta có thể sử…

