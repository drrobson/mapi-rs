use proc_macro::TokenStream;
use quote::{format_ident, quote};
use syn::{
    braced,
    parse::{Parse, ParseStream},
    parse_macro_input,
    punctuated::{Pair, Punctuated},
    token::Comma,
    Abi, Expr, ExprLit, FnArg, ForeignItemFn, Ident, Lit, LitStr, Meta, MetaNameValue, Pat,
    PatType, Result, ReturnType,
};

struct DelayLoadAttr {
    pub name: LitStr,
}

impl Parse for DelayLoadAttr {
    fn parse(input: ParseStream) -> Result<Self> {
        let meta: Meta = input.parse()?;
        match meta {
            Meta::NameValue(MetaNameValue {
                path,
                value:
                    Expr::Lit(ExprLit {
                        lit: Lit::Str(name),
                        ..
                    }),
                ..
            }) if path.get_ident().map(Ident::to_string).as_deref() == Some("name") => {
                Ok(DelayLoadAttr { name: name.clone() })
            }
            _ => Err(input.error(r#"expected #[delay_load(name = "...")]"#)),
        }
    }
}

struct ExternDecl {
    pub abi: LitStr,
    pub ident: Ident,
    pub inputs: Punctuated<FnArg, Comma>,
    pub output: ReturnType,
}

impl Parse for ExternDecl {
    fn parse(input: ParseStream) -> Result<Self> {
        let abi: Abi = input.parse()?;
        let abi = abi
            .name
            .ok_or_else(|| input.error(r#"expected "system" or "cdecl""#))?;

        let content;
        braced!(content in input);
        let foreign_item: ForeignItemFn = content.parse()?;

        Ok(ExternDecl {
            abi,
            ident: foreign_item.sig.ident,
            inputs: foreign_item.sig.inputs,
            output: foreign_item.sig.output,
        })
    }
}

/// Implement a delay load helper for the foreign function declaration in an extern block.
#[proc_macro_attribute]
pub fn delay_load(attr: TokenStream, input: TokenStream) -> TokenStream {
    let attr = parse_macro_input!(attr as DelayLoadAttr);
    let ast = parse_macro_input!(input as ExternDecl);
    impl_delay_load(&attr, &ast)
}

fn impl_delay_load(attr: &DelayLoadAttr, ast: &ExternDecl) -> TokenStream {
    let dll = &attr.name;
    let abi = &ast.abi;
    let name = &ast.ident;
    let inputs = &ast.inputs;
    let output = &ast.output;

    let mut args_size = quote! { 0 };
    let mut forward_args: Punctuated<Box<Pat>, Comma> = Punctuated::new();
    for pair in inputs.pairs() {
        match pair {
            Pair::Punctuated(FnArg::Typed(PatType { pat, ty, .. }), comma) => {
                forward_args.push_value(pat.clone());
                forward_args.push_punct(*comma);
                args_size = quote! { #args_size + mem::size_of::<#ty>() };
            }
            Pair::End(FnArg::Typed(PatType { pat, ty, .. })) => {
                forward_args.push_value(pat.clone());
                args_size = quote! { #args_size + mem::size_of::<#ty>() };
            }
            _ => panic!("should not have a receiver/self argument"),
        }
    }

    let func_type = format_ident!("PFN{}", name);
    let proc_name = LitStr::new(&format!("{name}"), name.span());
    let missing_export = LitStr::new(
        &format!("{name} is not exported from {}", dll.value()),
        name.span(),
    );

    let gen = quote! {
        unsafe fn #name(#inputs) #output {
            use std::{mem, sync::OnceLock};
            use ::windows_core::*;

            let mut proc_name: Vec<_> = #proc_name.bytes().collect();
            #[cfg(target_pointer_width = "32")]
            {
                const ARG_SIZE: usize = #args_size;
                proc_name.extend(format!("@{ARG_SIZE}").bytes());
            }
            proc_name.push(0);
            let proc_name = PCSTR::from_raw(proc_name.as_ptr());

            type #func_type = unsafe extern #abi fn(#inputs) #output;
            static EXPORT: OnceLock<#func_type> = OnceLock::new();

            (EXPORT.get_or_init(|| {
                use ::windows::Win32::System::LibraryLoader::*;

                unsafe {
                    let module = crate::get_mapi_module();
                    mem::transmute(GetProcAddress(module, proc_name).expect(#missing_export))
                }
            }))(#forward_args)
        }
    };

    gen.into()
}
