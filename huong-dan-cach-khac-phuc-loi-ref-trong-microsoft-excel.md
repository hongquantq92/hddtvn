Hướng dẫn cách khắc phục lỗi #REF! trong Microsoft Excel
========================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-cach-khac-phuc-loi-ref-trong-microsoft-excel/)Lỗi #REF! rất phổ biến khi sử dụng công thức trong bảng tính Excel, nhất là với hàm Vlookup hay Index. Lỗi này khá nghiêm trọng và trong nhiều trường hợp việc sửa lỗi này khiến bạn mất nhiều thời gian và công sức. Bài viết sẽ chỉ ra những nguyên nhân cơ bản gây …
-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Lỗi #REF! rất phổ biến khi sử dụng công thức trong bảng tính Excel, nhất là với hàm Vlookup hay Index. Lỗi này khá nghiêm trọng và trong nhiều trường hợp việc sửa lỗi này khiến bạn mất nhiều thời gian và công sức. Bài viết sẽ chỉ ra những nguyên nhân cơ bản gây ra lỗi #REF! và cách khắc phục nó giùm bạn.**


[![Hướng dẫn cách khắc phục lỗi #REF! trong Microsoft Excel](https://hddtvn.com/wp-content/uploads/2021/01/AJe50xb.png "Hướng dẫn cách khắc phục lỗi #REF! trong Microsoft Excel")](https://hddtvn.com/wp-content/uploads/2021/01/AJe50xb.png)


### 1. Lỗi #REF là gì?


REF là tên viết tắt của từ Reference, tức là Tham chiếu. Do đó trong mọi trường hợp xuất hiện lỗi này đều hiểu là lỗi tham chiếu.


Một số nguyên nhân chủ yếu gây ra lỗi #REF!:




* Tham chiếu tới 1 file đang chưa được mở nên không tham chiếu được.

* Đối tượng cần tham chiếu không có trong vùng tham chiếu.

* Đối tượng cần tham chiếu bị mất bởi việc xóa cột, xóa hàng chứa đối tượng đó, do bị ghi đè bởi việc copy paste



### 2. Khắc phục lỗi #REF do vùng tham chiếu bị xóa


Lỗi #REF! thường gặp khi bạn sử dụng hàm có chứa dữ liệu tham chiếu hoặc liên kết nhưng không thể tìm thấy do bị xóa.  Để xử lý trường hợp này, chúng ta cần cập nhật lại hàm, tham chiếu lại dữ liệu hoặc liên kết lại dữ liệu.


Ví dụ trong trường hợp này thì tham chiếu tại ô D14 đã bị xóa đi do đó kết quả sẽ trả về lỗi #REF!.


![](https://hddtvn.com/wp-content/uploads/2021/01/nwTRAfC.png)


Lúc này, ta cần cập nhật lại tham chiếu của ô để ô trả về kết quả bình thường.


![](https://hddtvn.com/wp-content/uploads/2021/01/F1Fgkrj.png)


### 3. Khắc phục lỗi #REF! do tham chiếu không có sẵn, vượt quá phạm vi


Việc tham chiếu vượt quá phạm vi cũng là nguyên nhân thường xuyên gây ra lỗi #REF!. Ví dụ trong trường hợp ở hình dưới, phạm vi của đối số **col\_index\_num** trong hàm VLOOKUP chỉ tối đa là 3. Vì vậy, khi ta để đối số là 4 thì sẽ vượt quá phạm vi tìm kiếm trong vùng **D1:F11** là 3 cột nên hàm sẽ trả về lỗi #REF!.


![](https://hddtvn.com/wp-content/uploads/2021/01/65fegiV.png)


Để giải quyết trường hợp này, chúng ta cần kiểm tra lại hàm, phạm vi tham chiếu của các đối số. Sau đó sửa lại cho chính xác là hàm sẽ hoạt động bình thường. Như trong trường hợp này thì ta cần sửa đối số **col\_index\_num** thành 2 do kết quả cần tìm tại cột E.


![](https://hddtvn.com/wp-content/uploads/2021/01/CcG2F3P.png)


### 4. Khắc phục lỗi #REF! do tham chiếu tới 1 file đang đóng


Thường xảy ra khi sử dụng hàm Indirect để tham chiếu tới 1 đối tượng bên ngoài workbook đang làm việc và workbook được tham chiếu đang đóng.


Để khắc phục trường hợp này thì ta cần mở Workbook được tham chiếu hoặc thay đổi phương pháp tham chiếu (không sử dụng hàm Indirect mà thay bằng hàm khác ít gặp lỗi hơn).


*Như vậy, bài viết trên đã hướng dẫn các bạn cách khắc phục lỗi #REF! trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreLỗi #REF! rất phổ biến khi sử dụng công thức trong bảng tính Excel, nhất…

