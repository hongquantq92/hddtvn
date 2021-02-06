Cách dùng hàm DATEDIF để tính thâm niên công tác
================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-dung-ham-datedif-de-tinh-tham-nien-cong-tac/)Khi tính lương theo hệ số lương, thâm niên công tác là một trong những yếu tố quan trọng. Cùng hddtvn tìm hiểu cách tính thâm niên công tác nhanh chóng và chính xác dưới đây. Sử dụng hàm DATEDIF để tính thâm niên công tác Để tính thâm niên công tác trên bảng tính …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Khi tính lương theo hệ số lương, thâm niên công tác là một trong những yếu tố quan trọng. Cùng hddtvn tìm hiểu cách tính thâm niên công tác nhanh chóng và chính xác dưới đây.**


![Hướng dẫn cách tính thâm niên công tác bằng Excel](https://hddtvn.com/wp-content/uploads/2021/01/s-cach-tinh-ngay-nghi-hang-nam-theo-tham-nien-cong-tac-2.png "Hướng dẫn cách tính thâm niên công tác bằng Excel")


Sử dụng hàm DATEDIF để tính thâm niên công tác
----------------------------------------------


Để tính thâm niên công tác trên bảng tính Excel, ta sử dụng hàm DATEDIF. Hàm này khi sử dụng sẽ trả về giá trị là số ngày (hoặc số năm, số tháng) giữa hai khoảng thời gian. Cú pháp của hàm DATEDIF như sau:


**DATEDIF(Firstdate, Enddate, Option)**


Trong đó:


– **Firstdate** là ngày bắt đầu của khoảng thời gian cần tính toán.


– **Enddate** là ngày kết thúc của khoảng thời gian cần tính toán.


– **Option**: quyết định kết quả tính toán sẽ trả về theo số ngày, tháng hay năm. Gồm có các tùy chọn sau:




* “d”: trả về số ngày giữa hai khoảng thời gian.

* “m”: trả về số tháng giữa hai khoảng thời gian (chỉ lấy phần nguyên).

* “y”: trả về số năm giữa hai khoảng thời gian (chỉ lấy phần nguyên).

* “yd”: trả về số ngày lẻ của năm giữa hai khoảng thời gian.

* “ym”: trả về số tháng lẻ của năm giữa hai khoảng thời gian.

* “md”: trả về số ngày lẻ của tháng giữa hai khoảng thời gian (số ngày chưa tròn tháng).



Hướng dẫn tính thâm niên công tác
---------------------------------


Ví dụ: ta cần tính thâm niên công tác của các nhân viên dưới bảng tính sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/jnMZVHY.png)


Tại ô **C3** của bảng tính, bạn nhập công thức tính như sau, sau đó ấn **Enter**:


**=DATEDIF(B3,TODAY(),"y")&" năm "&DATEDIF(B3,TODAY(),"ym")&" tháng "&DATEDIF(B3,TODAY(),"md")&" ngày"**


![Hướng dẫn cách tính thâm niên công tác bằng Excel](https://hddtvn.com/wp-content/uploads/2021/01/W3MKRG8.png "Hướng dẫn cách tính thâm niên công tác bằng Excel")


Trong công thức trên:


– **C3**: là ô đầu tiên trong cột “Thâm niên công tác” tương ứng với thâm niên của người thứ nhất trong bảng tính.


– **B3**: là ô chứa giá trị “Ngày vào công ty” của nhân viên thứ nhất trong bảng tính.


Sau khi ấn **ENTER**kết quả trả về sẽ chi tiết đến ngày, tháng, năm như sau:


![Hướng dẫn cách tính thâm niên công tác bằng Excel](https://hddtvn.com/wp-content/uploads/2021/01/IBqbxoG.png "Hướng dẫn cách tính thâm niên công tác bằng Excel")


Sau đó bạn sao chép công thức để tính thâm niên cho toàn bộ nhân viên có trong bảng tính. Kết quả nhận được như bảng dưới.


![Hướng dẫn cách tính thâm niên công tác bằng Excel](https://hddtvn.com/wp-content/uploads/2021/01/pwz7Sqc.png "Hướng dẫn cách tính thâm niên công tác bằng Excel")


Trên đây là cách tính thâm niên công tác nhanh chóng và dễ dàng trên Excel. Mời bạn đọc tham khảo và áp dụng. Chúc các bạn thành công.


moreKhi tính lương theo hệ số lương, thâm niên công tác là một trong những…

