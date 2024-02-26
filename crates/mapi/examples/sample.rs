use outlook_mapi::Session;

fn main() {
    println!("Trying to logon...");
    let _session = Session::new(true).expect("should be able to init and logon to MAPI");
    println!("Created the session...");
}
