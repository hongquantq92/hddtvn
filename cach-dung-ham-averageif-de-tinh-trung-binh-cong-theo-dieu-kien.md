Cách dùng hàm AVERAGEIF để tính trung bình cộng theo điều kiện
==============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-dung-ham-averageif-de-tinh-trung-binh-cong-theo-dieu-kien/)Thường khi tính trung bình cộng các số liệu thì bạn sẽ sử dụng tới hàm AVERAGE. Tuy nhiên, hầu như việc tính giá trị trung bình cộng trên Excel sẽ kèm theo nhiều điều kiện khác nhau. Và như vậy, bạn cần đến hàm AVERAGEIF để tính trung bình cộng có điều kiện. AVERAGEIF …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Thường khi tính trung bình cộng các số liệu thì bạn sẽ sử dụng tới hàm AVERAGE. Tuy nhiên, hầu như việc tính giá trị trung bình cộng trên Excel sẽ kèm theo nhiều điều kiện khác nhau. Và như vậy, bạn cần đến hàm AVERAGEIF để tính trung bình cộng có điều kiện. AVERAGEIF sẽ tiến hành tính trung bình cộng với những vùng dữ liệu được khoanh vùng có trong bảng, kèm theo đó là điều kiện cho trước mà người dùng thiết lập. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm AVERAGEIF trong Excel.**


![](https://hddtvn.com/wp-content/uploads/2021/01/avegeif.png)


### 1. Cấu trúc hàm AVERAGEIF


Cú pháp hàm: **=AVERAGEIF(range; criteria; [average\_range])**


*Trong đó:*




* **Range**: đối số bắt buộc, là một hoặc nhiều ô để tính giá trị trung bình. Bao gồm các số hoặc tên, mảng hoặc tham chiếu có chứa số.

* **Criteria**: đối số bắt buộc, là điều kiện để các ô được tính trung bình cộng.

* **Average\_range**: đối số tùy chọn, là tập hợp các ô để tính giá trị trung bình.



*Lưu ý:*




* Các ô trong **Range** có chứa TRUE hoặc FALSE hoặc là ô trống, hàm AVERAGE IF sẽ bỏ qua ô đó.

* Nếu **Range** là giá trị trống hoặc dạng văn bản, AVERAGEIF sẽ trả về giá trị lỗi #DIV0! .

* Nếu một ô trong **Average\_range** bị bỏ trống, hàm AVERAGEIF sẽ xem ô đó như giá trị 0.

* Nếu không có ô nào trong phạm vi tính đáp ứng điều kiện, hàm AVERAGEIF sẽ trả về giá trị lỗi #DIV/0! .

* Có thể dùng các ký tự đại diện, dấu chấm hỏi (?) và dấu sao (*) trong điều kiện. Dấu chấm hỏi sẽ khớp với bất kỳ ký tự đơn nào. Dấu sao sẽ khớp với bất kỳ chuỗi ký tự nào. Nếu muốn tìm một dấu chấm hỏi hay dấu sao thực sự, hãy gõ dấu ngã (~) trước ký tự.



### 2. Cách sử dụng hàm AVERAGEIF


Ví dụ ta có bảng dữ liệu như hình dưới, yêu cầu tính tiền hàng trung bình của các mặt hàng là Đá.


![](https://hddtvn.com/wp-content/uploads/2021/01/OMEg0eD.png)


Đầu tiên ta cần tạo ô điều kiện cho các mặt hàng là Đá


![](https://hddtvn.com/wp-content/uploads/2021/01/88Eljv8.png)


Ta có công thức tính trung bình cộng tiền hàng của các mặt hàng là Đá như sau:


**=AVERAGEIF(D2:D11;D13;F2:F11)**


*Trong đó:*




* **D2:D11** là vùng tên các loại hàng

* **D13** là điều kiện cần tính trung bình

* **F2:F11** là vùng chứa kết quả cần tính trung bình



![](https://hddtvn.com/wp-content/uploads/2021/01/GHJuJKk.png)


Hoặc các bạn cũng có thể thay ô điều kiện D13 bằng “Đá”


**=AVERAGEIF(D2:D11;”Đá”;F2:F11)**


![](https://hddtvn.com/wp-content/uploads/2021/01/dHQiyhW.png)


Trường hợp nếu cần tính trung bình cộng của các mặt hàng Đất và Đá, ta có thể dùng dấu * để thay thế cho các ký tự phía sau chữ Đ như sau:


**=AVERAGEIF(D2:D11;”Đ*”;F2:F11)**


![](https://hddtvn.com/wp-content/uploads/2021/01/fFODsfm.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm AVERAGEIF để tính trung bình cộng theo điều kiện. Chúc các bạn thành công!*


moreThường khi tính trung bình cộng các số liệu thì bạn sẽ sử dụng tới…

