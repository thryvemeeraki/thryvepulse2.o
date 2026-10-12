const {json,parseBody,requireEnv,AMOUNT_PAISE}=require('./_shared');
exports.handler=async function(event){
 if(event.httpMethod==='OPTIONS')return json(204,{}); if(event.httpMethod!=='POST')return json(405,{error:'Use POST to create a payment order.'});
 try{const b=parseBody(event);const name=String(b.name||'').trim(),email=String(b.email||'').trim().toLowerCase(),phone=String(b.phone||'').trim(),category=String(b.category||'');
 if(!name||!/^\S+@\S+\.\S+$/.test(email)||!phone||!['teen','adult'].includes(category))return json(400,{error:'Please complete all registration fields correctly.'});
 const keyId=requireEnv('RAZORPAY_KEY_ID'),secret=requireEnv('RAZORPAY_KEY_SECRET');const auth=Buffer.from(keyId+':'+secret).toString('base64');
 const r=await fetch('https://api.razorpay.com/v1/orders',{method:'POST',headers:{Authorization:'Basic '+auth,'Content-Type':'application/json'},body:JSON.stringify({amount:AMOUNT_PAISE,currency:'INR',receipt:'tp_'+Date.now(),notes:{name,email,phone,category,product:'Thryve Pulse Assessment'}})});const data=await r.json().catch(()=>({}));
 if(!r.ok||!data.id){console.error('Razorpay create order failed',r.status,JSON.stringify(data));return json(502,{error:'Razorpay could not create the order. Check matching Test/Live API credentials in Netlify.'});}return json(200,{keyId,orderId:data.id,amount:data.amount,currency:data.currency});
 }catch(e){console.error('create-order error:',e.message);return json(500,{error:e.message||'Unable to start payment.'});}
};
