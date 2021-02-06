Cách sử dụng hàm CONCATENATE để nối chuỗi trong Excel
=====================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-concatenate-de-noi-chuoi-trong-excel/)Như đã biết, dữ liệu trong bảng tính Excel không phải lúc nào cũng có cấu trúc theo nhu cầu của bạn. Bạn có thể muốn chia nội dung của một ô thành nhiều ô. Nhưng ngược lại, nếu cần kết hợp dữ liệu từ hai cột trở lên vào một cột đơn thì bạn …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Như đã biết, dữ liệu trong bảng tính Excel không phải lúc nào cũng có cấu trúc theo nhu cầu của bạn. Bạn có thể muốn chia nội dung của một ô thành nhiều ô. Nhưng ngược lại, nếu cần kết hợp dữ liệu từ hai cột trở lên vào một cột đơn thì bạn cần sử dụng hàm CONCATENATE. Hàm CONCATENATE được thiết kế để nối các đoạn văn bản khác nhau lại hoặc kết hợp các giá trị từ một vài ô vào một ô trong Excel.**


### 1. Cấu trúc hàm CONCATENATE


Cú pháp hàm: **=CONCATENATE(text1;[text2];…)**


*Trong đó:*




* text1 là đối số bắt buộc, là chuỗi đầu tiên cần ghép nối nó có thể là giá trị văn bản, số hoặc tham chiếu ô.

* text2 là đối số tùy chọn, là chuỗi thứ hai cần ghép nối có thể có tối đa 255 mục (8192 ký tự).



*Lưu ý:*




* Hàm CONCATENATE yêu cầu ít nhất một đối số **text** để làm việc.

* Các ký tự, ngày tháng… cần để trong dấu ngoặc kép ” “.

* Trong một công thức CONCATENATE, bạn có thể nối ghép lên tới 255 chuỗi, tổng cộng là 8,192 ký tự.

* Kết quả của hàm CONCATENATE luôn luôn là một chuỗi văn bản, ngay cả khi tất cả các giá trị nguồn là các con số.

* Nếu ít nhất một trong các đối số của hàm CONCATENATE không hợp lệ, công thức sẽ trả về lỗi #VALUE!.



### 2. Cách sử dụng hàm CONCATENATE


Ví dụ ta cần nối chuỗi hddtvn Nơi Kế Toán Kết Nối. Áp dụng cấu trúc hàm bên trên ta có công thức như sau:


**=CONCATENATE(D3,E3,F3,G3,H3,I3)**


![](https://hddtvn.com/wp-content/uploads/2021/01/5RudI2o.png)


Tuy nhiên nếu thực hiện theo cách trên thì bạn sẽ cần nhập từng đối số vào. Nếu như có quá nhiều đối số thì cách này không khả thi một chút nào. Để nhanh hơn, các bạn nhập hàm CONCATENATE. Sau đó quét mảng **D3:I3**. Tiếp theo các bạn bôi đen đối số **D3:I3** trong hàm rồi nhấn F9.


![](https://hddtvn.com/wp-content/uploads/2021/01/NdX8t4K.png)


Sau khi nhấn F9, chuỗi ký tự được tham chiếu của mảng D3:I3 sẽ hiện ra. Bây giờ các bạn chỉ cần xóa bỏ dấu  và dấu  ở đầu vào cuối mảng đi rồi nhấn **Enter** là được.


![](https://hddtvn.com/wp-content/uploads/2021/01/GhqcOkT.png)


Kết quả ta sẽ thu được chuỗi ký tự đã được nối vào với nhau một cách nhanh chóng.


![](https://hddtvn.com/wp-content/uploads/2021/01/TtVSK71.png)


Trường hợp muốn thêm dấu cách vào giữa mỗi chữ. Các bạn vẫn nhập hàm CONCATENATE rồi quét mảng D3:I3. Sau đó các bạn thêm chuỗi **&” “** vào đằng sau mảng D3:I3. Tiếp theo các bạn bôi đen toàn bộ đối số của hàm rồi nhấn F9.


![](https://hddtvn.com/wp-content/uploads/2021/01/vW7oeKM.png)


Sau khi nhấn F9, chuỗi ký tự được tham chiếu của mảng D3:I3 sẽ hiện ra kèm theo dấu cách ở đằng sau. Bây giờ các bạn chỉ cần xóa bỏ dấu  và dấu  ở đầu vào cuối mảng đi rồi nhấn **Enter** là được.


![](https://hddtvn.com/wp-content/uploads/2021/01/SotzQDi.png)


Kết quả ta sẽ thu được chuỗi ký tự đã được nối vào với nhau một cách nhanh chóng.


![](https://hddtvn.com/wp-content/uploads/2021/01/2qoSSiz.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm CONCATENATE để nối chuỗi trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreNhư đã biết, dữ liệu trong bảng tính Excel không phải lúc nào cũng có…

