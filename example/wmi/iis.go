package main

import (
	"fmt"
	"unsafe"

	"github.com/astaxie/beego/logs"

	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func init() {
	logs.EnableFuncCallDepth(true)
	logs.SetLogFuncCallDepth(3)
	logs.SetLevel(7)
}

func main() {
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	unknown, _ := oleutil.CreateObject("WbemScripting.SWbemLocator")
	defer unknown.Release()

	wmi, _ := unknown.QueryInterface(ole.IID_IDispatch)
	defer wmi.Release()

	// service is a SWbemServices
	serviceRaw, err := oleutil.CallMethod(wmi, "ConnectServer", "127.0.0.1", "root\\WebAdministration")
	if err != nil {
		fmt.Println(err)
	}
	service := serviceRaw.ToIDispatch()
	defer service.Release()

	// result is a SWBemObjectSet
	ssl2Raw, err := oleutil.CallMethod(service, "Get", "SSLBinding2")
	if err != nil {
		fmt.Println(ssl2Raw, err)
	}
	ssl2 := ssl2Raw.ToIDispatch()
	defer ssl2.Release()

	site, err := oleutil.CallMethod(ssl2, "Create", "*", 517, "3db569d8e9c788d0323008d5b72a56abc774a12a", "MY", "www.rango.com")
	if err != nil {
		logs.Debug(site, err)
		return
	}

	bindElementClsRaw, err := oleutil.CallMethod(service, "Get", "BindingElement")
	if err != nil {
		fmt.Println(err)
		return
	}
	bindElementCls := bindElementClsRaw.ToIDispatch()
	defer bindElementCls.Release()

	// 创建新的binding element
	bindElementRaw, err := oleutil.CallMethod(bindElementCls, "SpawnInstance_")
	if err != nil {
		fmt.Println(err)
		return
	}
	bindElementNew := bindElementRaw.ToIDispatch()
	defer bindElementNew.Release()

	_, err = bindElementNew.PutProperty("BindingInformation", "*:517:www.rango.com")
	_, err = bindElementNew.PutProperty("Protocol", "https")
	info, err := bindElementNew.GetProperty("BindingInformation")
	logs.Debug(info.ToString())
	if err != nil {
		fmt.Println(err)
		return
	}

	bindElementNew2 := bindElementRaw.ToIDispatch()
	defer bindElementNew2.Release()

	info, err = bindElementNew2.GetProperty("BindingInformation")
	logs.Debug(info.ToString())
	if err != nil {
		fmt.Println(err)
		return
	}

	defaultSiteRaw, err := oleutil.CallMethod(service, "Get", "Site.Name='default'") //
	if err != nil {
		fmt.Println(err)
		return
	}
	defaultSite := defaultSiteRaw.ToIDispatch()
	defer defaultSite.Release()

	// BindingElement 没有可用instance
	// bindingEles, err := oleutil.CallMethod(bindElementCls, "Instances_")
	// fmt.Println(bindingEles)
	// count, err := oleInt64(bindingEles.ToIDispatch(), "Count")
	// fmt.Println("网站Bindings实例", count)

	// _, err = oleutil.CallMethod(bindingEles.ToIDispatch(), "Item", "Bindings", bindElementNew)
	// if err != nil {
	// 	fmt.Println("Set Add", err)
	// 	return
	// }

	siteBindingsRaw, err := defaultSite.GetProperty("Bindings")
	if err != nil {
		fmt.Println(err)
		return
	}

	// 添加新的绑定信息
	AppendNewBinding(siteBindingsRaw, bindElementRaw)
	// *siteBindingsRaw = *newSiteBindings

	_, err = defaultSite.PutProperty("Bindings", siteBindingsRaw)
	if err != nil {
		fmt.Printf("PutProperty %v\n", err)
		return
	}
	_, err = defaultSite.CallMethod("Put_")
	if err != nil {
		fmt.Printf("PUT_ error %v\n", err)
		return
	}

	fmt.Println("Set Up SSL Certificate Successfully")
}

func oleInt64(item *ole.IDispatch, prop string) (int64, error) {
	v, err := oleutil.GetProperty(item, prop)
	if err != nil {
		return 0, err
	}
	i := int64(v.Val)
	return i, nil
}

// AppendNewBinding 追加新的SSL绑定
func AppendNewBinding(src *ole.VARIANT, dst *ole.VARIANT) *ole.VARIANT {
	siteBindingsConversion := src.ToArray()
	fmt.Printf("%p, %p\n", siteBindingsConversion.Array, dst)
	totalElements, _ := siteBindingsConversion.TotalElements(0)
	fmt.Println(totalElements)

	var bound ole.SafeArrayBound
	bound.LowerBound = 0
	bound.Elements = uint32(totalElements) + 1
	if err := siteBindingsConversion.Redim(&bound); err != nil {
		logs.Error(err)
		return nil
	}

	// totalElements, _ = siteBindingsConversion.TotalElements(0)
	// logs.Debug("after redim", totalElements)
	// siteBindings := siteBindingsConversion.ToValueArray()
	// logs.Debug(siteBindings)

	if err := siteBindingsConversion.PutElement(totalElements, uintptr(unsafe.Pointer(dst))); err != nil {
		logs.Error(err)
		return nil
	}

	// totalElements, _ = siteBindingsConversion.TotalElements(0)
	// logs.Debug("after redim", totalElements)

	// siteBindings = siteBindingsConversion.ToValueArray()
	// logs.Debug(siteBindings)

	return nil

	// siteBindings = append(siteBindings, dst)
	// fmt.Printf("after append: %p, %p\n", siteBindings[0], siteBindings[1])

	// var newSiteBindings ole.VARIANT
	// ole.VariantInit(&newSiteBindings)

	// // var testdata = []string{"Ads", "good"}
	// // testArray := ole.SafeArrayFromStringSlice(testdata)
	// // fmt.Println(testArray)
	// newSiteBindingsArray := ole.SafeArrayFromVariantSlice(siteBindings)
	// // fmt.Println(newSiteBindingsArray, int64(uintptr(newSiteBindingsArray.Data)))
	// newSiteBindings = ole.NewVariant(ole.VT_ARRAY|ole.VT_VARIANT, int64(uintptr(unsafe.Pointer(newSiteBindingsArray))))
	// fmt.Println(newSiteBindings)

	// siteBindingsConversion2 := newSiteBindings.ToArray()
	// totalElements2, _ := siteBindingsConversion2.TotalElements(0)
	// fmt.Println(totalElements2)

	// for i, v := range siteBindingsConversion2.ToValueArray() {
	// 	fmt.Println(i, v)
	// 	certHash, err := v.(*ole.IDispatch).GetProperty("Protocol")
	// 	if err != nil {
	// 		fmt.Println(err)
	// 	}
	// 	fmt.Println(certHash.ToString())
	// }
	// // siteBindingsConversion.Array = siteBindings

	// return &newSiteBindings
}

// ModifySSLBinding 修改网站SSL绑定
func ModifySSLBinding() {

}

// EnumSite 枚举网站
func EnumSite(defaultSite *ole.IDispatch) {
	siteInstsRaw, err := oleutil.CallMethod(defaultSite, "Instances_")
	if err != nil {
		fmt.Println(err)
		return
	}
	siteInsts := siteInstsRaw.ToIDispatch()
	defer siteInsts.Release()

	count, err := oleInt64(siteInsts, "Count")
	fmt.Println("网站实例", count)

	for i := int64(0); i < count; i++ {
		itemRaw, err := oleutil.CallMethod(siteInsts, "ItemIndex", i)
		if err != nil {
			fmt.Println("枚举失败", err)
			break
		}
		item := itemRaw.ToIDispatch()
		defer item.Release()

		port, err := item.GetProperty("Name")
		if err != nil {
			fmt.Println("获取网站名称失败", err)
			break
		}
		if port.ToString() == "Default Web Site" {

		}
		fmt.Println(port.ToString())
	}

}
