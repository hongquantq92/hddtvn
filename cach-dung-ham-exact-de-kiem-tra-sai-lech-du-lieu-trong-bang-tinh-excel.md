Cách dùng hàm EXACT để kiểm tra sai lệch dữ liệu trong bảng tính Excel
======================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-dung-ham-exact-de-kiem-tra-sai-lech-du-lieu-trong-bang-tinh-excel/)Khi sử dụng Excel để phân tích dữ liệu, độ chính xác là yếu tố quan trọng nhất. Các công thức của Excel luôn luôn đúng, nhưng kết quả có thể sai do trong hệ thống có dữ liệu lỗi. Trong trường hợp này, biện pháp duy nhất là kiểm tra dữ liệu và độ …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Khi sử dụng Excel để phân tích dữ liệu, độ chính xác là yếu tố quan trọng nhất. Các công thức của Excel luôn luôn đúng, nhưng kết quả có thể sai do trong hệ thống có dữ liệu lỗi. Trong trường hợp này, biện pháp duy nhất là kiểm tra dữ liệu và độ chính xác. Bài viết viết đây sẽ hướng dẫn các bạn những cách để so sánh hai ô dữ liệu trong Excel một cách đơn giản nhất.**


### 1. Cấu trúc hàm EXACT


Cú pháp hàm: **=EXACT(text1; text2)**


*Trong đó:*


Text1: đối số bắt buộc, là chuỗi văn bản thứ nhất.


Text2: đối số bắt buộc, là chuỗi văn bản thứ hai.


*Lưu ý:*




* Hàm EXACT sẽ trả về TRUE nếu chúng hoàn toàn giống nhau, FALSE nếu khác.

* Hàm EXACT phân biệt chữ hoa, chữ thường nhưng bỏ qua khác biệt về định dạng.

* Nếu không muốn phân biệt chữ hoa và chữ thường, bạn có thể dùng công thức =A=B



### 2. Cách sử dụng hàm EXACT


Ví dụ ta cần so sánh dữ liệu Text 1 và Text 2 như bảng sau đây xem có giống nhau không.


[![](https://hddtvn.com/wp-content/uploads/2021/01/hNDItoZ.png)](https://hddtvn.com/wp-content/uploads/2021/01/hNDItoZ.png)


Áp dụng cấu trúc hàm EXACT như trên ta có công thức so sánh tại ô C2 như sau:


**=EXACT(A2;B2)**


![](https://hddtvn.com/wp-content/uploads/2021/01/Amax4Et.png)


Như ta thấy hàm sẽ trả về kết quả là FALSE do Text 1 có chữ K hoa còn Text 2 có chữ k thường. Hàm EXACT sẽ phân biệt cả chữ hoa và chữ thường. Nếu không muốn phân biệt chữ hoa và chữ thường ta có thể sử dụng công thức =A=B để so sánh. Ta sẽ có công thức so sánh tại ô D2 như sau:


**=A2=B2**


![](https://hddtvn.com/wp-content/uploads/2021/01/NZjLPZO.png)


Ta có thể thấy công thức =A=B sẽ trả về kết quả là TRUE do không phân biệt chữ hoa và chữ thường. Ta có thể thấy kết quả cũng sẽ tương tự đối với ô C3 và D3.


Sao chép công thức của ô C2 và D2 cho các ô còn lại ta sẽ thu được kết quả so sánh cả cột Text 1 và Text 2 như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/mqyZTNE.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm EXACT và công thức =A=B để so sánh dữ liệu của 2 ô trong Excel. Chúc các bạn thành công!*


moreKhi sử dụng Excel để phân tích dữ liệu, độ chính xác là yếu tố…

