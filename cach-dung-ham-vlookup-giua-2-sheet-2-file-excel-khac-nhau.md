Cách dùng hàm VLOOKUP giữa 2 Sheet, 2 file Excel khác nhau
==========================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-dung-ham-vlookup-giua-2-sheet-2-file-excel-khac-nhau/)VLOOKUP là hàm khá thông dụng và được sử dụng thường xuyên trên bảng tính Excel. Thông thường hàm VLOOKUP hay áp dụng trên cùng một file cùng một bảng tính nhất định. Nhưng có trường hợp vùng điều kiện và vùng bạn cần truy xuất dữ liệu lại thuộc vào 2 sheet khác nhau, …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**VLOOKUP là hàm khá thông dụng và được sử dụng thường xuyên trên bảng tính Excel. Thông thường hàm VLOOKUP hay áp dụng trên cùng một file cùng một bảng tính nhất định. Nhưng có trường hợp vùng điều kiện và vùng bạn cần truy xuất dữ liệu lại thuộc vào 2 sheet khác nhau, thậm chí xa hơn có thể là 2 file excel khác nhau. Khi đó bạn sẽ làm gì?Hãy đọc bài viết sau để biết cách dùng hàm VLOOKUP giữa 2 Sheet, 2 file Excel khác nhau nhé.**


### 1. Sử dụng hàm VLOOKUP giữa 2 sheet Excel khác nhau


Ví dụ ta có sheet Sản phẩm chứa bảng dữ liệu và sheet Thuế suất chứa bảng thuế suất của từng loại hàng. Yêu cần cần tham chiếu thuế suất của từng mặt hàng giữa hai bảng với nhau. Để thực hiện, đầu tiên các bạn cần để con trỏ chuột tại ô **G2** của sheet **Sản phẩm** sau đó nhập công thức như sau:


**=VLOOKUP(D2;**


![](https://hddtvn.com/wp-content/uploads/2021/01/12-2.png)


Tiếp theo, các bạn chuyển sang sheet **Thuế suất** rồi nhập tiếp:


**=VLOOKUP(D2;'Thuế suất'!$B$2:$C$5;2;FALSE)**


![](https://hddtvn.com/wp-content/uploads/2021/01/11-2.png)


Sau khi nhập xong, các bạn nhấn **Enter** rồi sao chép công thức cho các ô còn lại. Kết quả ta sẽ thu được là cột thuế suất đã được tham chiếu thành công.


![](https://hddtvn.com/wp-content/uploads/2021/01/13-2.png)


### 2. Sử dụng hàm VLOOKUP giữa 2 file Excel khác nhau


Ví dụ ta có 2 file Excel, một file chứa bảng dữ liệu và một file chứa bảng thuế suất của từng loại hàng. Yêu cần cần tham chiếu thuế suất của từng mặt hàng giữa hai file với nhau. Để thực hiện, đầu tiên các bạn cần để con trỏ chuột tại ô **G2** của file chứa bảng dữ liệu, sau đó nhập công thức như sau:


**=VLOOKUP(D2;**


![Cách dùng hàm VLOOKUP giữa 2 Sheet, 2 file Excel khác nhau](https://hddtvn.com/wp-content/uploads/2021/01/14-1.png "Cách dùng hàm VLOOKUP giữa 2 Sheet, 2 file Excel khác nhau")


Tiếp theo, các bạn chuyển sang file chứa bảng thuế suất rồi nhập tiếp:


**=VLOOKUP(D2;'[Book2]Thuế suất'!B$2:C$5;2;FALSE)**


![Cách dùng hàm Vlookup giữa 2 Sheet, 2 file Excel khác nhau](https://hddtvn.com/wp-content/uploads/2021/01/15-1.png "Cách dùng hàm Vlookup giữa 2 Sheet, 2 file Excel khác nhau")


Sau khi nhập công thức xong, các bạn nhấn **Enter** rồi sao chép công thức cho các ô còn lại. Kết quả ta sẽ thu được là cột thuế suất đã được tham chiếu thành công.


![Cách dùng hàm Vlookup giữa 2 Sheet, 2 file Excel khác nhau](https://hddtvn.com/wp-content/uploads/2021/01/16-1.png "Cách dùng hàm Vlookup giữa 2 Sheet, 2 file Excel khác nhau")


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm Vlookup giữa 2 sheet và 2 file Excel khác nhau. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


#### 


moreVLOOKUP là hàm khá thông dụng và được sử dụng thường xuyên trên bảng tính…

