using FireSharp;
using FireSharp.Config;
using FireSharp.Interfaces;

namespace HSCTSV
{
    class FirebaseConnect
    {
        private static FirebaseConnect instance;
        public static FirebaseConnect Instance
        {
            get { if (instance == null) instance = new FirebaseConnect(); return instance; }
            private set { instance = value; }
        }
        public IFirebaseClient client;
        public FirebaseConnect()
        {
            IFirebaseConfig config = new FirebaseConfig()
            {
                AuthSecret = "rs9XLzp48e24Ds3HX6QSGV6wFCVcaVAxMOqoT1V9",
                BasePath = "https://hsctsv-54653.firebaseio.com/"
            };
            client = new FirebaseClient(config);
        }
    }
}
