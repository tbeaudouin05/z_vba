Attribute VB_Name = "c_generate_all_countries"

Sub generate_hk()

Application.ScreenUpdating = False

Call show_all

'generate hk

Call filter_country("hk")
        
Call seller_CN_index_and_overviews
    
Call adapt_template_for_TW

Sheets("Seller_CN_index").calculate

Call generate_all

Application.ScreenUpdating = True

End Sub

Sub generate_sg()

Application.ScreenUpdating = False

Call show_all

'generate hk

Call filter_country("sg")
        
Call seller_CN_index_and_overviews
    
Call adapt_template_for_TW

Sheets("Seller_CN_index").calculate

Call generate_all

Application.ScreenUpdating = True

End Sub

Sub generate_tw()

Application.ScreenUpdating = False

Call show_all

'generate hk

Call filter_country("tw")
        
Call seller_CN_index_and_overviews
    
Call adapt_template_for_TW

Sheets("Seller_CN_index").calculate

Call generate_all

Application.ScreenUpdating = True

End Sub

Sub generate_my()

Application.ScreenUpdating = False

Call show_all

'generate hk

Call filter_country("my")
        
Call seller_CN_index_and_overviews
    
Call adapt_template_for_TW

Sheets("Seller_CN_index").calculate

Call generate_all

Application.ScreenUpdating = True

End Sub

Sub generate_all_countries()

Application.ScreenUpdating = False

Call show_all

'generate hk

Call filter_country("hk")
        
Call seller_CN_index_and_overviews
    
Call adapt_template_for_TW

Sheets("Seller_CN_index").calculate

Call generate_all


' generate sg

Call show_all

Call filter_country("sg")
        
Call seller_CN_index_and_overviews
    
Call adapt_template_for_TW

Sheets("Seller_CN_index").calculate

Call generate_all

        
' generate tw

Call show_all

Call filter_country("tw")
        
Call seller_CN_index_and_overviews
    
Call adapt_template_for_TW

Sheets("Seller_CN_index").calculate

Call generate_all


' generate my

Call show_all

Call filter_country("my")
        
Call seller_CN_index_and_overviews
    
Call adapt_template_for_TW

Sheets("Seller_CN_index").calculate

Call generate_all

Application.ScreenUpdating = True

End Sub
